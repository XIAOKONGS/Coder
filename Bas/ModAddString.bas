Attribute VB_Name = "ModAddString"
Option Explicit
'*************************************************************************
'**ģ �� ����ModAddString
'**˵    �����Զ�����ַ�����Ŀ���ַ�����β
'**�� �� �ˣ�XIAOKONGS
'**��    �ڣ�2017��6��15��
'**��    ע: ��Ȩ���� XIAOKONGS 2017
'*************************************************************************

Public Function AddStrToStr(ByVal Str1 As String, ByVal Str2 As String) As String
    '�Զ�����ַ�����Ŀ���ַ�����β
    If LCase(Right(Str1, Len(Str2))) = LCase(Str2) Then
        AddStrToStr = Str1
    Else
        AddStrToStr = Str1 & Str2
    End If
End Function
