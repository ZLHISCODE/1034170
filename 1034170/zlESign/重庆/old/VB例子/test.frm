VERSION 5.00
Object = "{1D774E06-D3E3-48F7-842C-8C39CC14D299}#1.0#0"; "CQCATO~1.OCX"
Begin VB.Form Form1 
   Caption         =   "签名验证测试程序(VB)"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10125
   StartUpPosition =   3  '窗口缺省
   Begin CQCATOOLKITSLib.CQCAToolkits CQCAToolkits1 
      Height          =   0
      Left            =   7560
      TabIndex        =   12
      Top             =   2880
      Width           =   0
      _Version        =   65536
      _ExtentX        =   0
      _ExtentY        =   0
      _StockProps     =   0
   End
   Begin VB.CommandButton Command8 
      Caption         =   "指定证书"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "证书序列号"
      Height          =   495
      Left            =   8040
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "证书主题信息"
      Height          =   495
      Left            =   8040
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "证书主题名称"
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "查看证书"
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "选择证书"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "验证"
      Height          =   495
      Left            =   8040
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "签名"
      Height          =   495
      Left            =   8040
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   2295
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "签名生成数据:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "签名原始数据:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim sign As New CERTSIGNVERIFYLib.CertSign
'Dim verify As New CERTSIGNVERIFYLib.CertVerify
Public cqcakits As Object
Public str As String
Public strCertName As String


Private Sub Command1_Click()
    Text2.Text = CQCAToolkits1.SignData(strCertName, Text1.Text)
End Sub

Private Sub Command2_Click()
    verifyvalue = CQCAToolkits1.VerifyData(Text1.Text, Text2.Text)
    If verifyvalue = "valid signature" Then
        MsgBox "数据验证成功!", vbOKOnly, "检验数据"
    Else
         MsgBox "数据验证失败!", vbOKOnly, "校验数据"
    End If
End Sub

Private Sub Command3_Click()

str1 = CQCAToolkits1.SelectCert("My", 1)
If str1 = "error" Then
    'MsgBox "未选择证书"
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Exit Sub
End If
str = str1
strCertName = CQCAToolkits1.GetSelectCert
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
End Sub

Private Sub Command4_Click()
    CQCAToolkits1.ShowCert str
End Sub
Private Sub Command5_Click()
    MsgBox CQCAToolkits1.GetSelectCert
End Sub

Private Sub Command6_Click()
MsgBox CQCAToolkits1.GetCertSubject

End Sub

Private Sub Command7_Click()
    MsgBox CQCAToolkits1.GetCertNumber

End Sub

Private Sub Command8_Click()
Dim certstr As String
certstr = InputBox("请输入证书序列号!")
str1 = CQCAToolkits1.SelectCert(certstr, 0) '4A78CE277DC3E8D4D709AF75EA3CFA53
If str1 = "error" Then
    'MsgBox "未选择证书"
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Exit Sub
End If
str = str1
strCertName = CQCAToolkits1.GetSelectCert
Command1.Enabled = True
Command2.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
End Sub

Private Sub Form_Load()

On Error GoTo err1
Set cqcakits = CreateObject("CQCATOOLKITS.CQCAToolkitsCtrl.1") '动态创建对象

Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Exit Sub
err1:
    Command1.Enabled = False
    Command2.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command3.Enabled = False
    MsgBox "控件未注册"
End Sub

