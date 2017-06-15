Attribute VB_Name = "ModGetFileNameInPath"
Option Explicit
'*************************************************************************
'**模 块 名：ModGetFileNameInPath
'**说    明：从一个全路径字符串中找到文件名,支持UNC路径
'**创 建 人：XIAOKONGS
'**日    期：2017年6月15日
'**备    注: 版权所有 XIAOKONGS 2017
'*************************************************************************

Public Function GetFileNameInPath(ByVal FullPathName As String, Optional ByVal NoExtName As Boolean = False) As String
    '从指定全路径中找到文件名
    'FullPathName指定全路径
    '返回值:包含的文件名
    Dim I As Long, J As Long
    Dim FileName As String, FileNameNoExt As String
    
    FullPathName = Trim(FullPathName)
    I = InStrRev(FullPathName, "\")
    J = Len(FullPathName)
    If I = 0 Then
        I = InStrRev(FullPathName, "/")
        J = Len(FullPathName)
    End If
    If I = 0 Then Exit Function
    
    FileName = Mid(FullPathName, I + 1, J - I)
    I = InStrRev(FileName, ".")
    J = Len(FileName)
    If I = 0 Then Exit Function
    
    FileNameNoExt = Mid(FileName, 1, I - 1)
    If NoExtName = True Then
        GetFileNameInPath = FileNameNoExt
    Else
        GetFileNameInPath = FileName
    End If
End Function

Public Function GetDirInPath(ByVal FullPathName As String, Optional ByVal pFull As Boolean = True) As String
    '从指定全路径中找到目录或路径
    'FullPathName指定全路径
    'pFull为True时返回路径,为False时返回目录名
    '返回值:pFull所指定的值
    Dim I As Long, J As Long
    Dim tPath As String, tDir As String
    
    FullPathName = Trim(FullPathName)
    I = InStrRev(FullPathName, "\")
    J = Len(FullPathName)
    If I = 0 Then
        I = InStrRev(FullPathName, "/")
        J = Len(FullPathName)
    End If
    If I = 0 Then Exit Function
    
    tPath = Mid(FullPathName, 1, I - 1)
    I = InStrRev(tPath, "\")
    J = Len(tPath)
    If I = 0 Then
        I = InStrRev(tPath, "/")
        J = Len(tPath)
    End If
    If I = 0 Then Exit Function
    
    tDir = Mid(tPath, I + 1, J - I)
    If pFull = True Then
        GetDirInPath = tPath
    Else
        GetDirInPath = tDir
    End If
End Function
