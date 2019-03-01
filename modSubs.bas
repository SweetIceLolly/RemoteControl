Attribute VB_Name = "modSubs"
Public Const EncryptNumber = 13                 '密匙 (*^-^*)

'==========================================================================================
'注意：以下开关在编译前必须全部设定为True，否则程序运行可能不正常。
Public Const bCompress = True                   '调试用开关：是否压缩数据
Public Const bEncrypt = True                    '调试用开关：是否加密明文数据
Public Const bKillTasks = True                  '调试用开关：是否结束进程
Public Const bKeepSending = True                '调试用开关：是否连续发送图片
Public Const bMouseWheelHook = True             '调试用开关：是否拦截鼠标滚轮消息
Public Const bBlockInput = True                 '调试用开关：是否允许禁用鼠标键盘
Public Const bDeleteFiles = True                '调试用开关：是否允许删除文件

'==========================================================================================
'以下内容为配置文件里面配置的内容
Public bAutoStart As Boolean                    '是否开机自动启动
Public bTrayWhenStart As Boolean                '启动后是否最小化到托盘区
Public bHideMode As Boolean                     '是否以隐藏模式运行
Public sUserPassword As String                  '用户密码
Public bUseUserPassword As Boolean              '是否使用用户密码
Public bAutoRecord As Boolean                   '是否自动记录IP和密码
Public bHideWhenStart As Boolean                '是否自动启动后以后台模式运行
Public bNoExit As Boolean                       '是否禁止退出
Public bExitWhenClose As Boolean                '关闭时是否退出 【True：关闭时退出软件 False：关闭时最小化到托盘区】

'==========================================================================================
'以下为远控时的状态
Public IsRemoteControl As Boolean               '是否为远程控制状态 【True:远程控制模式 False：文件传输模式】
Public IsControlling As Boolean                 '是否控制对方鼠标键盘
Public IsWorking As Boolean                     '当前是否有上传任务或者下载任务正在进行
Public AutoResize As Boolean                    '是否拉伸画面
Public ScreenScaleW As Double                   '缩小屏幕和原始屏幕的宽度比例尺
Public ScreenScaleH As Double                   '缩小屏幕和原始屏幕的高度比例尺

Public Function Encrypt(strEncrypt As String) As String     '加密（可以进行逆运算的哦~）
    Dim tmp As String
    '==============================
    If bEncrypt = False Then                    '根据调试开关数值判断是否进行加解密
        Encrypt = strEncrypt
        Exit Function
    End If
    For i = 1 To Len(strEncrypt)
        tmp = tmp & Chr(Asc(Left(Mid(strEncrypt, i), 1)) Xor EncryptNumber)   '哈哈，Xor加密简单粗暴还可以逆运算
    Next i
    Encrypt = tmp
End Function

Public Function Decrypt(strDecrypt As String) As String     '解密
    If bEncrypt = False Then                                    '根据调试开关数值判断是否进行加解密
        Decrypt = strDecrypt
        Exit Function
    End If
    Decrypt = Encrypt(strDecrypt)                               '再调用一次加密进行逆运算就是解密啦~
End Function

Public Sub wsSendData(TargetWinsock As Winsock, DataToSend As String)         '发送数据过程
    On Error Resume Next                                    '让我来拯救世界！！
    TargetWinsock.SendData Encrypt(DataToSend & "{S}{END}")
End Sub

Public Sub LoadConfig()                                     '加载配置文件过程
    On Error Resume Next
    Dim tmpString As String                                 '每行内容
    Dim tmpSplit() As String                                '切割后的字符串缓存
    Dim sIP() As String                                     'IP和对应的密码的缓存
    '----------------------------
    Open App.Path + "\Config.ini" For Input As #1               '打开文件
        If Err.Number <> 0 Then                                 '若无法打开文件则按照程序的默认值来配置
            bAutoStart = False
            bTrayWhenStart = False
            bHideMode = False
            sUserPassword = "123456"
            '----------------------------
            Call SaveConfig                                         '自动生成配置文件
            Exit Sub                                                '退出过程
        End If
        '----------------------------
        Do While Not EOF(1)
            Line Input #1, tmpString                            '逐行读取内容
            tmpSplit = Split(tmpString, "=")                    '分割
            Select Case LCase(tmpSplit(0))                      '分类讨论 【转成小写是为了适应性强啦。。】
                Case "autostart"                                    '是否开机自动启动
                    bAutoStart = CBool(tmpSplit(1))                 '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bAutoStart = False
                    End If
                    
                Case "traywhenstart"                                '启动后是否最小化到托盘区
                    bTrayWhenStart = CBool(tmpSplit(1))             '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bTrayWhenStart = False
                    End If
                
                Case "hidemode"                                     '是否以隐藏模式运行
                    bHideMode = CBool(tmpSplit(1))                  '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bHideMode = False
                    End If
                
                Case "userpassword"                                 '用户密码
                    sUserPassword = Decrypt(tmpSplit(1))                '读取的时候解码
                    
                Case "autoresize"                                   '是否自动缩放图像
                    AutoResize = CBool(tmpSplit(1))                 '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        AutoResize = False
                    End If
                    
                Case "useuserpassword"                              '是否使用用户密码
                    bUseUserPassword = CBool(tmpSplit(1))           '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bUseUserPassword = False
                    End If
                
                Case "autorecord"                                   '是否自动记录密码
                    bAutoRecord = CBool(tmpSplit(1))                '将数据类型转换为布尔类
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bAutoRecord = False
                    End If
                
                Case "hidewhenstart"                                '是否自动启动后以后台模式运行
                    bHideWhenStart = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bHideWhenStart = False
                    End If
                            
                Case "noexit"                                       '是否禁止退出
                    bNoExit = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bNoExit = False
                    End If
                    
                Case "exitwhenclose"                                '关闭时是否退出
                    bExitWhenClose = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '如果读取到无效数据就用默认的数据
                        bExitWhenClose = False
                    End If
                
                Case Else
                    If InStr(tmpSplit(0), "[") = 0 Then             '如果不是标记就记录到列表里
                        sIP = Split(tmpSplit(0), "大腿")                '分割IP和密码 【哈哈。“大腿”是个小彩蛋哟~用“大腿”作为分隔符。。亏我想得出】
                        frmMain.comIP.AddItem sIP(0)                    '添加IP
                        frmMain.lstPassword.AddItem Decrypt(sIP(1))     '添加密码
                    End If
                    
            End Select
        Loop
    Close #1
End Sub

Public Sub SaveConfig()                                     '保存配置文件过程
    Open App.Path + "\Config.ini" For Output As #1              '打开文件
        Print #1, "[Config]"                                        '文件第一行，虽然加不加这行都无关紧要，但是加上这个好看~
        Print #1, "AutoStart=" & bAutoStart                         '是否开机自启
        Print #1, "TrayWhenStart=" & bTrayWhenStart                 '启动后是否最小化到托盘区
        Print #1, "HideMode=" & bHideMode                           '是否以隐藏模式运行
        Print #1, "UserPassword=" & Encrypt(sUserPassword)          '用户密码 【为了安全，将密码加密了再保存】
        Print #1, "AutoResize=" & AutoResize                        '是否拉伸图像
        Print #1, "UseUserPassword=" & bUseUserPassword             '是否使用用户密码
        Print #1, "AutoRecord=" & bAutoRecord                       '是否自动记录IP
        Print #1, "HideWhenStart=" & bHideWhenStart                 '是否自动启动后以后台模式运行
        Print #1, "NoExit=" & bNoExit                               '是否禁止退出
        Print #1, "ExitWhenClose=" & bExitWhenClose                 '关闭时是否退出
        Print #1, "[List]"                                          '又是文件里面区分的标志，同样是为了美观，不过去掉也无所谓~ ╮(╯_╰)╭
        For i = 0 To frmMain.comIP.ListCount - 1                    '列表里的IP和对应的密码 【同样是为了安全，将密码加密了再保存】
            Print #1, frmMain.comIP.List(i) & "大腿" & Encrypt(frmMain.lstPassword.List(i))
        Next i
    Close #1
End Sub
