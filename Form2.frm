VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Caption         =   "产品认证页面 (在线认证)"
   ClientHeight    =   11415
   ClientLeft      =   120
   ClientTop       =   435
   ClientWidth     =   16815
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11415
   ScaleWidth      =   16815
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton refresh 
      BackColor       =   &H00C0FFC0&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15600
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser browser 
      Height          =   10455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   16575
      ExtentX         =   29236
      ExtentY         =   18441
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "认证参数："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   15135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
browser.Navigate Form1.url
    If Left(Form1.jhm, 2) = "0x" Then
        Label1.ForeColor = RGB(255, 120, 0)
        Label1.Caption = "激活码：" & Form1.jhm & " / 产品代码：" & Form1.cpdm & "  ||  ■ 测试版(～￣▽￣)～"
    Else
        Label1.Caption = "激活码：" & Form1.jhm & " / 产品代码：" & Form1.cpdm & "  ||  感谢您选择并支持 正版 氖光(TM) 产品！"
    End If
End Sub

Private Sub refresh_Click()
browser.refresh

Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 1 "  '历史记录
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 2 "  'cookies
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 8 "  '临时文件
'Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 16 "
'Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 32 "
'Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255 "
'Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351 "   '清空浏览器数据
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell "RunDll32.exe InetCpl.cpl,ClearMyTracksByProcess 255 "
End Sub
