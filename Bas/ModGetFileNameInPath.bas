Attribute VB_Name = "ModGetFileNameInPath"
Option Explicit
'*************************************************************************
'**ģ �� ����ModGetFileNameInPath
'**˵    ������һ��ȫ·���ַ������ҵ��ļ���,֧��UNC·��
'**�� �� �ˣ�XIAOKONGS
'**��    �ڣ�2017��6��15��
'**��    ע: ��Ȩ���� XIAOKONGS 2017
'*************************************************************************

Public Function GetFileNameInPath(ByVal FullPathName As String, Optional ByVal NoExtName As Boolean = False) As String
    '��ָ��ȫ·�����ҵ��ļ���
    'FullPathNameָ��ȫ·��
    '����ֵ:�������ļ���
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
    '��ָ��ȫ·�����ҵ�Ŀ¼��·��
    'FullPathNameָ��ȫ·��
    'pFullΪTrueʱ����·��,ΪFalseʱ����Ŀ¼��
    '����ֵ:pFull��ָ����ֵ
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
