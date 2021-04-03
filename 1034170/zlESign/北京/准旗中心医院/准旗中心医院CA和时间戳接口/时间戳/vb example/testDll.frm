VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7500
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "获取时间戳"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "验证"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   2055
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   2640
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "sha1"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "申请"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

    Dim obj As Object
    Set obj = CreateObject("tsaMiddleware.UtilUdp")     '调用dll控件
    
    Dim data1 As String     '需要加盖时间戳的数据
    Dim data2 As String     'hash算法，暂时只支持sha1
    Dim result As String    '返回结果
    
    data1 = Text1.Text
    data2 = Text2.Text
    
    result = obj.sendTimestamp(data1, data2)
    
    Text3.Text = result
End Sub

Private Sub Command2_Click()

    Dim obj As Object
    Set obj = CreateObject("tsaMiddleware.UtilUdp")

    Dim data1 As String     '需要验证时间戳的数据
    Dim data2 As String     'hash算法，暂时只支持sha1
    Dim result As String    '返回结果
    
    data1 = Text1.Text
    data2 = Text2.Text
    
    result = obj.verifyTimestamp(data1, data2)
    
    Text3.Text = result
End Sub

Private Sub Command3_Click()

    Dim obj As Object
    Set obj = CreateObject("tsaMiddleware.UtilUdp")

    Dim data1 As String     '需要获得时间戳信息的数据
    Dim data2 As String     'hash算法，暂时只支持sha1
    Dim result As String    '返回结果
    
    data1 = Text1.Text
    data2 = Text2.Text
    
    result = obj.getTimestampInfo(data1, data2)
    
    Text3.Text = result
End Sub

