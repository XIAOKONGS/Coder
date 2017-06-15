Attribute VB_Name = "ModAddString"
Option Explicit
'*************************************************************************
'**模 块 名：ModAddString
'**说    明：自动添加字符串到目标字符串结尾
'**创 建 人：XIAOKONGS
'**日    期：2017年6月15日
'**备    注: 版权所有 XIAOKONGS 2017
'*************************************************************************

Public Function AddStrToStr(ByVal Str1 As String, ByVal Str2 As String) As String
    '自动添加字符串到目标字符串结尾
    If LCase(Right(Str1, Len(Str2))) = LCase(Str2) Then
        AddStrToStr = Str1
    Else
        AddStrToStr = Str1 & Str2
    End If
End Function
