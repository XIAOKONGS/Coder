VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件更新 DEMO"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5460
   FillStyle       =   4  'Upward Diagonal
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5460
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5400
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   5340
      ScaleWidth      =   5400
      TabIndex        =   5
      Top             =   0
      Width           =   5460
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   7080
      Width           =   930
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "检查更新"
      Height          =   345
      Left            =   4440
      TabIndex        =   3
      Top             =   5520
      Width           =   930
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   540
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   180
      Width           =   4965
   End
   Begin VB.Label labProgress 
      Caption         =   "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5610
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'XIAOKONGS 文件下载类 DEMO
'
'输入下载地址,目标路径,即开始下载.
'下载进度事件中可得到下载速度,进度等详细信息.
'
'版权所有 XIAOKONGS 2017

Dim WithEvents oDownload As CFileDownload
Attribute oDownload.VB_VarHelpID = -1

Private Sub oDownload_OnProgress(ByVal lProgress As Long, ByVal lMaxProgress As Long, ByVal lSpeed As Long, ByVal lStatusCode As Long, ByVal sStatusText As String)
    Dim I As Single
    
    If lProgress = 0 Or lMaxProgress = 0 Then Exit Sub
    
    I = lProgress / lMaxProgress * 30
    labProgress.Caption = String(I, "■")
    Me.Caption = "文件下载DEMO(" & Format(I, "0.00") & "%,速度 = " & lSpeed & "KB/S)"
End Sub

Private Sub cmdDownload_Click()

    If cmdDownload.Caption = "完成" Then End

    '設置佈局狀態
    cmdDownload.Caption = "更新中"
    cmdDownload.Enabled = False
    
    '开始下载
    '
    '注意StartDownloading过程是阻塞的.
    
    If oDownload.StartDownloading(txtURL.Text, AddStrToStr(App.Path, "\") & GetFileNameInPath(txtURL.Text)) Then
'        MsgBox "下载成功!", vbOKOnly Or vbInformation
        cmdDownload.Caption = "完成"
        cmdDownload.Enabled = True
        Me.Caption = "XIAOKONGS 文件更新 DEMO v" & App.Major & App.Revision
    Else
        MsgBox "下载失败!", vbOKOnly Or vbInformation
    End If
End Sub

Private Sub cmdCancel_Click()
    '取消下载.
    Call oDownload.AbortDownloading
End Sub

Private Sub Form_Load()
    Set oDownload = New CFileDownload
    
    Me.Caption = "XIAOKONGS 文件更新 DEMO v" & App.Major & App.Revision
    
'    txtURL.Text = "http://www.m5home.com/soft/vb_link.rar"                      'VB函数添加大师
    txtURL.Text = "https://github.com/XIAOKONGS/FILES/raw/master/Calling.exe"        '飞信
    labProgress.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call oDownload.AbortDownloading
    Set oDownload = Nothing
End Sub
