VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "产品认证助手(demo) v.2.0"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5535
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton jh 
      BackColor       =   &H0080C0FF&
      Caption         =   "激活认证"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "输入产品代码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
      Begin VB.TextBox cpdmt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "输入认证代码"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      Begin VB.TextBox t3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox t2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox t1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line3 
         X1              =   3000
         X2              =   3120
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   1080
         X2              =   1200
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Menu 菜单 
      Caption         =   "菜单"
      Begin VB.Menu 关于产品认证助手 
         Caption         =   "关于“产品认证助手(demo)”"
      End
      Begin VB.Menu fgx 
         Caption         =   "-"
      End
      Begin VB.Menu 检查更新 
         Caption         =   "检查更新"
      End
   End
   Begin VB.Menu 清空内容 
      Caption         =   "清空内容"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public jhm$
Public cpdm$
Public url$
Dim mrym$

Private Sub jh_Click()

If t1.Text = "" Or t2.Text = "" Or t3.Text = "" Or cpdmt.Text = "" Then
    MsgBox "输入信息不完全 → 请检查“认证代码”或“产品代码”！", vbCritical, "【×】错误！Σ(っ °Д °;)っ"
Else
    jhm = t1.Text & "-" & t2.Text & "-" & t3.Text
    cpdm = cpdmt.Text
    mrym = "index.html"
    url = "https://mosaik404.github.io/products-quali/demo/download/quali/" & cpdm & "/" & jhm & "/" & mrym
    MsgBox "即将打开认证页面。→ 本功能需要联网。", vbInformation, "【i】提示！(o゜▽゜)o☆"
    '用系统默认浏览器打开：
    'Shell "explorer.exe " & url
    
    '指定浏览器内核为IE11，防止显示错误
    CreateObject("wscript.shell").regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
    
    Form2.Show
    
End If
End Sub

Private Sub 关于氖光产品认证助手_Click()
MsgBox "氖光(TM) 产品认证助手 v.2.0 - 2024-04-19", vbInformation, "【i】提示！(o゜▽゜)o☆"
End Sub

Private Sub 检查更新_Click()
MsgBox "当前版本：v.2.0，打开网页检查更新。", vbInformation, "【i】提示！(o゜▽゜)o☆"
End Sub

Private Sub 清空内容_Click()
t1.Text = ""
t2.Text = ""
t3.Text = ""
cpdmt.Text = ""
End Sub
