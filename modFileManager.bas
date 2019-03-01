Attribute VB_Name = "modFileManager"
'文件浏览模块

'获取磁盘空间的API
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
'获取文件信息的API
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const MAX_PATH = 260

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA            '文件信息
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public isCopy As Boolean            'True = 复制 ；False = 剪切


Public Function SizeWithFormat(lBytes) As String        '字节数转换单位过程
    Select Case lBytes
        Case Is < 1024                           '<1024 - Byte
           SizeWithFormat = lBytes & " Byte"
        
        Case Is < 1024 ^ 2                      '<1024^2 - KB
            SizeWithFormat = Format(lBytes / 1024, "0.00") & " KB"
        
        Case Is < 1024 ^ 3                      '<1024^3 - MB
            SizeWithFormat = Format(lBytes / (1024 ^ 2), "0.00") & " MB"
        
        Case Is < 1024 ^ 4                      '<1024^4 - GB
            SizeWithFormat = Format(lBytes / (1024 ^ 3), "0.00") & " GB"
        
    End Select
End Function

Public Function GetFileSize(sFilePath As String) As String          '获取文件大小过程
    Dim lFile As Long
    Dim FileMsg As WIN32_FIND_DATA
    '=======================================
    lFile = FindFirstFile(sFilePath, FileMsg)
    lFileSize = FileMsg.nFileSizeHigh * 4294967295# + FileMsg.nFileSizeLow          '洛小羽の方法1
    If lFileSize < 0 Then
        lFileSize = 2147483647 + lFileSize                                          '洛小羽の方法2
        lFileSize = lFileSize + 2147483647 + 2
    End If
    GetFileSize = SizeWithFormat(lFileSize)
End Function


Public Sub MakeList(DirPath As String)          '生成目录文件表的过程
    '名称|类型|大小
    Dim tmp() As String                         '缓存字符串
    Dim FilePath As String                      '文件路径字符串
    On Error Resume Next
    '=================================
    frmMain.lstTemp.Clear
    frmMain.lstTemp.AddItem "...|上层目录|:N:"
    '=================================
    frmMain.Dir.Path = DirPath
    frmMain.Dir.Refresh
    If Err.Number <> 0 Then
        MakeRootList
        Exit Sub
    End If
    For i = 0 To frmMain.Dir.ListCount - 1
        tmp = Split(frmMain.Dir.List(i), "\")
        frmMain.lstTemp.AddItem tmp(UBound(tmp)) & "|目录|:N:"
    Next i
    '=================================
    frmMain.File.Path = DirPath
    frmMain.File.Refresh
    FilePath = IIf(Right(frmMain.File.Path, 1) = "\", frmMain.File.Path, frmMain.File.Path & "\")
    For i = 0 To frmMain.File.ListCount - 1
        tmp = Split(frmMain.File.List(i), "\")
        frmMain.lstTemp.AddItem tmp(UBound(tmp)) & "|文件|" & GetFileSize(FilePath & frmMain.File.List(i))
    Next i
End Sub

Public Sub MakeRootList()                       '生成磁盘根目录表的过程
    Dim tmpString As String
    Dim lSPC                         'Sectors Per Cluster        【每簇的扇区数】
    Dim lBPS                         'Bytes Per Sector           【每扇区的字节数】
    Dim lF                           'Number Of Free Clusters    【空闲簇的数量】
    Dim lT                           'Total Number Of Clusters   【簇的总数】
    Dim IsExists As Long             '磁盘是否有效
    '===============================
    frmMain.lstTemp.Clear
    frmMain.Drive.Refresh
    For i = 0 To frmMain.Drive.ListCount - 1
        tmpString = frmMain.Drive.List(i)        '盘符/名称
        IsExists = GetDiskFreeSpace(Left(tmpString, 2), lSPC, lBPS, lF, lT)          '获取硬盘空间
        tmpString = tmpString & "|（空闲:" & SizeWithFormat(lSPC * lBPS * lF)
        tmpString = tmpString & " 共:" & SizeWithFormat(lSPC * lBPS * lT) & "）"
        If IsExists <> 0 Then                    '如果能获取到磁盘信息说明磁盘有效
            frmMain.lstTemp.AddItem tmpString
        End If
    Next i
End Sub

Public Function IsPathExists(strPath As String, Optional bAcceptFiles As Boolean = False) As Boolean    '判断路径是否存在
    On Error Resume Next
    If Dir(strPath) = "" Then                                           '如果用Dir检测不到路径的话返回False
        IsPathExists = False
        Exit Function
    Else
        If bAcceptFiles = True Then
            IsPathExists = True
            Exit Function
        End If
    End If
    Open strPath For Binary As #1                                       '如果是文件的话同样是返回False
        If Err.Number = 0 Then
            If bAcceptFiles = False Then
                IsPathExists = False
            Else
                IsPathExists = True
            End If
            Close #1
            Exit Function
        End If
    Close #1
    IsPathExists = True                                                 '两层选拔都过关就返回True
End Function

