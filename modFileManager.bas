Attribute VB_Name = "modFileManager"
'�ļ����ģ��

'��ȡ���̿ռ��API
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
'��ȡ�ļ���Ϣ��API
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long

Public Const MAX_PATH = 260

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA            '�ļ���Ϣ
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

Public isCopy As Boolean            'True = ���� ��False = ����


Public Function SizeWithFormat(lBytes) As String        '�ֽ���ת����λ����
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

Public Function GetFileSize(sFilePath As String) As String          '��ȡ�ļ���С����
    Dim lFile As Long
    Dim FileMsg As WIN32_FIND_DATA
    '=======================================
    lFile = FindFirstFile(sFilePath, FileMsg)
    lFileSize = FileMsg.nFileSizeHigh * 4294967295# + FileMsg.nFileSizeLow          '��С��η���1
    If lFileSize < 0 Then
        lFileSize = 2147483647 + lFileSize                                          '��С��η���2
        lFileSize = lFileSize + 2147483647 + 2
    End If
    GetFileSize = SizeWithFormat(lFileSize)
End Function


Public Sub MakeList(DirPath As String)          '����Ŀ¼�ļ���Ĺ���
    '����|����|��С
    Dim tmp() As String                         '�����ַ���
    Dim FilePath As String                      '�ļ�·���ַ���
    On Error Resume Next
    '=================================
    frmMain.lstTemp.Clear
    frmMain.lstTemp.AddItem "...|�ϲ�Ŀ¼|:N:"
    '=================================
    frmMain.Dir.Path = DirPath
    frmMain.Dir.Refresh
    If Err.Number <> 0 Then
        MakeRootList
        Exit Sub
    End If
    For i = 0 To frmMain.Dir.ListCount - 1
        tmp = Split(frmMain.Dir.List(i), "\")
        frmMain.lstTemp.AddItem tmp(UBound(tmp)) & "|Ŀ¼|:N:"
    Next i
    '=================================
    frmMain.File.Path = DirPath
    frmMain.File.Refresh
    FilePath = IIf(Right(frmMain.File.Path, 1) = "\", frmMain.File.Path, frmMain.File.Path & "\")
    For i = 0 To frmMain.File.ListCount - 1
        tmp = Split(frmMain.File.List(i), "\")
        frmMain.lstTemp.AddItem tmp(UBound(tmp)) & "|�ļ�|" & GetFileSize(FilePath & frmMain.File.List(i))
    Next i
End Sub

Public Sub MakeRootList()                       '���ɴ��̸�Ŀ¼��Ĺ���
    Dim tmpString As String
    Dim lSPC                         'Sectors Per Cluster        ��ÿ�ص���������
    Dim lBPS                         'Bytes Per Sector           ��ÿ�������ֽ�����
    Dim lF                           'Number Of Free Clusters    �����дص�������
    Dim lT                           'Total Number Of Clusters   ���ص�������
    Dim IsExists As Long             '�����Ƿ���Ч
    '===============================
    frmMain.lstTemp.Clear
    frmMain.Drive.Refresh
    For i = 0 To frmMain.Drive.ListCount - 1
        tmpString = frmMain.Drive.List(i)        '�̷�/����
        IsExists = GetDiskFreeSpace(Left(tmpString, 2), lSPC, lBPS, lF, lT)          '��ȡӲ�̿ռ�
        tmpString = tmpString & "|������:" & SizeWithFormat(lSPC * lBPS * lF)
        tmpString = tmpString & " ��:" & SizeWithFormat(lSPC * lBPS * lT) & "��"
        If IsExists <> 0 Then                    '����ܻ�ȡ��������Ϣ˵��������Ч
            frmMain.lstTemp.AddItem tmpString
        End If
    Next i
End Sub

Public Function IsPathExists(strPath As String, Optional bAcceptFiles As Boolean = False) As Boolean    '�ж�·���Ƿ����
    On Error Resume Next
    If Dir(strPath) = "" Then                                           '�����Dir��ⲻ��·���Ļ�����False
        IsPathExists = False
        Exit Function
    Else
        If bAcceptFiles = True Then
            IsPathExists = True
            Exit Function
        End If
    End If
    Open strPath For Binary As #1                                       '������ļ��Ļ�ͬ���Ƿ���False
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
    IsPathExists = True                                                 '����ѡ�ζ����ؾͷ���True
End Function

