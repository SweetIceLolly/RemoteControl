VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "任务 (1/1)"
   ClientHeight    =   1344
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5508
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1344
   ScaleWidth      =   5508
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   840
   End
   Begin MSWinsockLib.Winsock wsFile 
      Left            =   3360
      Top             =   840
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5292
      _Version        =   786432
      _ExtentX        =   9334
      _ExtentY        =   444
      _StockProps     =   93
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdOpenLocalDir 
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   864
      Width           =   1092
      _Version        =   786432
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "中断任务"
      ForeColor       =   16777215
      BackColor       =   16416003
      Appearance      =   3
      DrawFocusRect   =   0   'False
      ImageGap        =   1
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2333 GB/s 2333 GB/2333 GB 100%"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2880
   End
   Begin VB.Label labFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "正在下载：C:\Windows\System32\I Love CXT.exe"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4164
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileList As String                                   '文件列表
Public IsUpload As Boolean                                  '是否为上传状态 【True:上传 False:下载】
Public CurrentSize As Long                                  '当前文件大小
Public CurrentFileIndex As Integer                          '当前的文件序号
Public TotalFiles As Integer                                '文件总数
Public DownloadPath As String                               '文件下载的路径
Public TotalSent As Integer                                 '文件发送数

Public lSent As Long                                        '已经发送的数据字节数
Public oldSent As Long                                      '上一秒发送的数据字节数
Public BytesPerSec As Long                                  '每秒数据字节数

Public lRec As Long                                         '接收的字节数
Public lOldRec As Long                                      '上一秒接收的字节数

Dim Temp() As Byte                                          '打开的文件的内容
Dim EachFile() As String                                    '每个文件的路径

Public Function LoadFile(sFileName As String) As Boolean    '读取文件过程
    On Error Resume Next
    '-----------------------------------
    LoadFile = True
    Open sFileName For Binary As #1                             '打开文件
        If Err.Number <> 0 Then                                 '文件不存在
            If sFileName <> "" Then                                 '如果不是Split导致的空文件名
                frmFileTransfer.AddLog "文件 " & sFileName & " 读取失败。将跳过上传该文件。"
            End If
            Close #1                                            '关闭文件
            LoadFile = False
            Exit Function
        End If
        '--------------------
        ReDim Temp(LOF(1))                                      '分配内存
        CurrentSize = LOF(1)
        If Err.Number <> 0 Then                                 '文件过大
            frmFileTransfer.AddLog "文件过大：" & sFileName & " 将跳过上传该文件。"
            LoadFile = False
            Exit Function
        End If
        If UBound(Temp) = 0 Then                                '如果读取到的字节数为0
            frmFileTransfer.AddLog "文件 " & sFileName & " 读取失败，该文件可能受保护或为空文件，将跳过上传该文件。"
            Close #1                                            '关闭文件
            LoadFile = False
            Exit Function
        End If
        Get #1, , Temp                                          '读取文件
    Close #1
    lSent = 0                                               '清空文件发送字节数
    oldSent = 0
End Function

Public Function NextFile() As Boolean                       '下一个文件过程
    CurrentFileIndex = CurrentFileIndex + 1                     '文件序号 + 1
    If CurrentFileIndex > TotalFiles Then                       '如果序号超出文件总数
        NextFile = False
        frmFileTransfer.AddLog "上传任务结束。共发送" & TotalSent & "个文件。"
        Unload Me
        Exit Function
    End If
    '---------------------------
    Do While LoadFile(EachFile(CurrentFileIndex)) = False   '如果读取文件错误就继续下一个文件
        CurrentFileIndex = CurrentFileIndex + 1                     '文件序号 + 1
        If CurrentFileIndex > TotalFiles Then                       '如果序号超出文件总数
            NextFile = False
            frmFileTransfer.AddLog "上传任务结束。共发送" & TotalSent & "个文件。"
            Exit Function
        End If
    Loop
    TotalSent = TotalSent + 1
    Me.labFile.Caption = "正在上传：" & EachFile(CurrentFileIndex)
    Me.Caption = "任务 (" & CurrentFileIndex + 1 & "/" & TotalFiles & ")"
    frmFileTransfer.AddLog "正在上传：" & EachFile(CurrentFileIndex)
    NextFile = True
End Function

Public Function SplitPath()                                 '分理出每个文件的路径过程
    EachFile = Split(FileList, vbCrLf)                          '分割
    TotalFiles = UBound(EachFile)                               '计算出文件总数
End Function

Public Function SendUploadMsg()                             '发送上传消息过程
    '发送上传文件请求 【请求头|文件路径|文件大小】
    wsSendData frmFileTransfer.wsMessage, "UPL|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsWorking Then                   '如果当前有正在进行中的任务
        Dim a
        a = MsgBox("确认关闭本窗口？当前进行的任务将中断！", 32 + vbYesNo, "确认")
        If a <> 6 Then                      '取消关闭
            Cancel = True
        Else                                '确认关闭
            frmFileTransfer.AddLog "文件上传中断。共发送" & TotalSent & "个文件。"
            wsSendData frmFileTransfer.wsMessage, "END"     '发送结束上传消息
            wsSendData frmFileTransfer.wsMessage, "REF"     '发送刷新请求
        End If
    End If
End Sub

Private Sub tmrRefresh_Timer()
    If IsUpload Then
        BytesPerSec = lSent - oldSent                           '计算每秒发送的字节数
        oldSent = lSent                                         '上一秒的字节数变成这一秒的
    Else
        BytesPerSec = lRec - lOldRec                            '计算每秒接收的字节数
        lOldRec = lRec                                          '上一秒的字节数变成这一秒的
    End If
End Sub

Private Sub wsFile_ConnectionRequest(ByVal requestID As Long)
    '接收到连接请求
    Me.wsFile.Close
    Me.wsFile.Accept requestID                              '建立连接
    If IsUpload Then                                        '判断是否为上传文件状态
        Me.wsFile.SendData "CNT"                                '连接建立确认消息
    Else
        Me.wsFile.SendData "DNT"                                '让对方准备发送
    End If
End Sub

Private Sub wsFile_DataArrival(ByVal bytesTotal As Long)
    Dim recTemp() As Byte               '接收到的消息的缓存
    Dim a As VbMsgBoxResult             'Msgbox 返回值
    '------------------------
    Me.wsFile.GetData recTemp
    '得到前三位
    Select Case Chr(recTemp(0)) & Chr(recTemp(1)) & Chr(recTemp(2))
        Case "TSN"                                          '文件重名确认
            a = MsgBox("文件重名：" & EachFile(CurrentFileIndex) & "，是否确认覆盖？" & vbCrLf & "点击【是】将覆盖文件，点击【否】将取消发送该文件。", 32 + vbYesNo, "文件覆盖确认")
            If a = 6 Then
                frmFileTransfer.AddLog "确认覆盖文件 " & EachFile(CurrentFileIndex)
                Me.wsFile.SendData "YES"                        '发送 确认覆盖 消息
            Else
                frmFileTransfer.AddLog "取消覆盖文件 " & EachFile(CurrentFileIndex)
                If NextFile = True Then
                    '如果成功打开下一个文件就上传下一个文件请求
                    wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
                Else
                    wsSendData frmFileTransfer.wsMessage, "END"     '打开失败就发送结束上传消息
                    wsSendData frmFileTransfer.wsMessage, "REF"     '发送刷新请求
                    IsWorking = False                               '当前没有任务
                    Unload Me
                End If
            End If
        
        Case "IRD"                                          '对方准备就绪
            Me.tmrRefresh.Enabled = True                        '激活统计计时器
            Me.wsFile.SendData Temp                             '开始砸数据 (╯‵□′)╯︵┻━┻
            
        Case "DFL"                                          '覆盖失败
            frmFileTransfer.AddLog "文件 " & EachFile(CurrentFileIndex) & " 上传失败。原因：覆盖文件失败"
            If NextFile = True Then
                '如果成功打开下一个文件就上传下一个文件请求
                wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
            Else
                wsSendData frmFileTransfer.wsMessage, "END"     '打开失败就发送结束上传消息
                wsSendData frmFileTransfer.wsMessage, "REF"     '发送刷新请求
                IsWorking = False                               '当前没有任务
                Unload Me
            End If
        
        Case "NXT"                                          '下一文件请求
            If NextFile = True Then
                '如果成功打开下一个文件就发送上传下一个文件请求
                wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
            Else
                wsSendData frmFileTransfer.wsMessage, "END"     '打开失败就发送结束上传消息
                wsSendData frmFileTransfer.wsMessage, "REF"     '发送刷新请求
                IsWorking = False                               '当前没有任务
                Unload Me
            End If
        
        Case "NNT"                                          '准备好接收下一个文件
            Me.wsFile.SendData "CNT"
        
        Case "DNT"
            Dim nFileTitle As String, nFileSize As String       '文件标题 文件大小
            Dim SplitTmp() As String                            '分割缓存
            Dim recString As String                             '把接收到的数据转换成字符串的变量
            Dim recTotal As Integer, recCurrent As Integer      '需要接受的文件总数 当前接收的文件序号
            '----------------------------------
            recString = StrConv(recTemp, vbUnicode)                 '将二进制数据转换成字符串
            '----------------------------------
            SplitTmp = Split(Replace(recString, "DNT|", ""), "|")   '去掉请求头
            nFileTitle = SplitTmp(0)                                '分离出文件标题
            nFileSize = SplitTmp(1)                                 '分离出文件大小
            recCurrent = SplitTmp(2)                                '分离出当前接收的文件序号
            recTotal = SplitTmp(3)                                  '分离出需要接受的文件总数
            Me.Caption = "任务 (" & recCurrent + 1 & "/" & recTotal & ")"   '更新标题
            lRec = 0                                                '清空接收字节数
            lOldRec = 0
            frmFileTransfer.AddLog "开始下载文件：" & nFileTitle
            Me.wsFile.Tag = nFileTitle & "|" & nFileSize            '赋值到控件的Tag
            Me.labFile.Caption = "正在下载：" & nFileTitle          '更新窗体状态
            '----------------------------------
            If IsPathExists(DownloadPath & nFileTitle, True) = True Then    '如果文件已存在
                a = MsgBox("检测到文件已存在：" & DownloadPath & nFileTitle & "，是否确认覆盖？" & vbCrLf & "点击【是】将覆盖文件，点击【否】将取消发送该文件。", 32 + vbYesNo, "文件覆盖确认")
                If a = 6 Then                                                   '确认覆盖
                    Kill DownloadPath & nFileTitle                                  '尝试删除文件
                    If Err.Number = 0 Then                                          '删除成功
                        frmFileTransfer.AddLog "覆盖文件 " & DownloadPath & nFileTitle
                        Open DownloadPath & nFileTitle For Binary As #2                 '打开文件
                        Me.wsFile.SendData "IRD"                                        '告诉对方准备就绪
                    Else                                                            '删除失败
                        frmFileTransfer.AddLog "覆盖文件 " & DownloadPath & nFileTitle & " 失败，将取消下载文件 " & nFileTitle
                        Me.wsFile.SendData "NXT"                                        '告诉对方下一文件
                    End If
                Else
                    frmFileTransfer.AddLog "取消覆盖文件 " & DownloadPath & nFileTitle
                    Me.wsFile.SendData "NXT"                                        '告诉对方下一文件
                End If
            Else                                                            '如果文件不存在
                Open DownloadPath & nFileTitle For Binary As #2                 '打开文件
                Me.wsFile.SendData "IRD"                                        '果树对方准备就绪
            End If
            
        Case "END"                                                      '下载完毕
            frmFileTransfer.AddLog "下载任务结束。共下载" & Split(StrConv(recTemp, vbUnicode), "|")(1) & "个文件。"
            IsWorking = False                                           '当前没有任务
            frmFileTransfer.cmdRefreshLocal_Click                       '刷新本地目录
            Unload Me
            
        Case Else                                                       '如果是其他数据
            '得到后三位
            '如果是【FIN】，即单个文件下载完毕
            If recTemp(UBound(recTemp) - 2) = 70 And recTemp(UBound(recTemp) - 1) = 73 And recTemp(UBound(recTemp)) = 78 Then
                frmFileTransfer.AddLog "文件 " & Split(Me.wsFile.Tag, "|")(0) & " 下载完毕。"
                Close #2                                                    '关闭文件
                Me.wsFile.SendData "NXT"                                    '告诉对方下一文件
            Else                                                        '如果仍然是其他数据 说明不是命令
                Put #2, LOF(2) + 1, recTemp                                 '写入文件
                lRec = lRec + UBound(recTemp)
                '-----------------------------------------
                '更新窗体状态
                Dim tmpShow As String                                       '标签上显示的内容的缓存
                Dim CurrentFileSize As Long                                 '当前文件大小
                CurrentFileSize = CLng(Split(Me.wsFile.Tag, "|")(1))                            '从Tag获取当前文件大小
                tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                    '每秒字节数
                tmpShow = tmpShow & " " & SizeWithFormat(lRec)                                  '已经发送的大小
                tmpShow = tmpShow & "/" & SizeWithFormat(Split(Me.wsFile.Tag, "|")(1))          '文件的大小
                tmpShow = tmpShow & " " & Format((lRec / CurrentFileSize) * 100, "0.00") & "%"  '计算百分比
                Me.ProgressBar.Value = (lRec / CurrentFileSize) * 100                           '计算任务完成的百分比
                Me.labStatus.Caption = tmpShow                                                  '显示到标签上
                Me.labStatus.Refresh                                                            '刷新标签
            End If
        
    End Select
End Sub

Private Sub wsFile_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If IsUpload Then
        lSent = lSent + bytesSent
        If lSent >= CurrentSize Then
            lSent = 0
            frmFileTransfer.AddLog "文件" & EachFile(CurrentFileIndex) & "上传成功。"
            Me.wsFile.SendData "FIN"
            Exit Sub
        End If
        '=======================================================
        '更新窗体状态
        Dim tmpShow As String                                                       '显示的文本缓存
        tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                '每秒字节数 = 现在发送的数据字节数 - 上一秒发送的数据字节数
        tmpShow = tmpShow & " " & SizeWithFormat(lSent)                             '已经发送的大小
        tmpShow = tmpShow & "/" & SizeWithFormat(CurrentSize)                       '文件的大小
        tmpShow = tmpShow & " " & Format(lSent / CurrentSize * 100, "0.00") & "%"   '计算百分比
        Me.ProgressBar.Value = lSent / CurrentSize * 100
        Me.labStatus.Caption = tmpShow                                              '显示到标签上
        Me.labStatus.Refresh
    End If
End Sub
