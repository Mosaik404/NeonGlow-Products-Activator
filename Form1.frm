VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��Ʒ��֤����(demo) v.2.0"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   5535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5535
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton jh 
      BackColor       =   &H0080C0FF&
      Caption         =   "������֤"
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
      Caption         =   "�����Ʒ����"
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
      Caption         =   "������֤����"
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
   Begin VB.Menu �˵� 
      Caption         =   "�˵�"
      Begin VB.Menu ���ڲ�Ʒ��֤���� 
         Caption         =   "���ڡ���Ʒ��֤����(demo)��"
      End
      Begin VB.Menu fgx 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
   End
   Begin VB.Menu ������� 
      Caption         =   "�������"
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
    MsgBox "������Ϣ����ȫ �� ���顰��֤���롱�򡰲�Ʒ���롱��", vbCritical, "���������󣡦�(�� �㧥 ��;)��"
Else
    jhm = t1.Text & "-" & t2.Text & "-" & t3.Text
    cpdm = cpdmt.Text
    mrym = "index.html"
    url = "https://mosaik404.github.io/products-quali/demo/download/quali/" & cpdm & "/" & jhm & "/" & mrym
    MsgBox "��������֤ҳ�档�� ��������Ҫ������", vbInformation, "��i����ʾ��(o�b���b)o��"
    '��ϵͳĬ��������򿪣�
    'Shell "explorer.exe " & url
    
    'ָ��������ں�ΪIE11����ֹ��ʾ����
    CreateObject("wscript.shell").regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
    
    Form2.Show
    
End If
End Sub

Private Sub �����ʹ��Ʒ��֤����_Click()
MsgBox "�ʹ�(TM) ��Ʒ��֤���� v.2.0 - 2024-04-19", vbInformation, "��i����ʾ��(o�b���b)o��"
End Sub

Private Sub ������_Click()
MsgBox "��ǰ�汾��v.2.0������ҳ�����¡�", vbInformation, "��i����ʾ��(o�b���b)o��"
End Sub

Private Sub �������_Click()
t1.Text = ""
t2.Text = ""
t3.Text = ""
cpdmt.Text = ""
End Sub
