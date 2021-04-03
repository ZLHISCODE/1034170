VERSION 5.00
Begin VB.Form frmIdentify重庆 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保病人身份识别"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIdentify重庆.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "白内障摘除术"
      Height          =   240
      Index           =   3
      Left            =   4110
      TabIndex        =   11
      Top             =   2670
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   0
      TabIndex        =   9
      Top             =   1350
      Width           =   6660
   End
   Begin VB.Frame Frame2 
      Height          =   1785
      Left            =   3570
      TabIndex        =   10
      Top             =   1260
      Width           =   30
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Left            =   -210
      TabIndex        =   6
      Top             =   2985
      Width           =   6660
   End
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "急诊抢救"
      Height          =   240
      Index           =   2
      Left            =   4110
      TabIndex        =   5
      Top             =   2310
      Width           =   1275
   End
   Begin VB.OptionButton opt类别 
      BackColor       =   &H8000000A&
      Caption         =   "特殊病门诊"
      Height          =   240
      Index           =   1
      Left            =   4110
      TabIndex        =   4
      Top             =   1950
      Width           =   1515
   End
   Begin VB.OptionButton opt类别 
      Caption         =   "普通门诊"
      Height          =   240
      Index           =   0
      Left            =   4110
      TabIndex        =   3
      Top             =   1590
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.TextBox txtEdit 
      Height          =   420
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1860
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   2355
      TabIndex        =   7
      Top             =   3210
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   3990
      TabIndex        =   8
      Top             =   3210
      Width           =   1305
   End
   Begin VB.Image Image2 
      Height          =   1005
      Left            =   270
      Picture         =   "frmIdentify重庆.frx":030A
      Stretch         =   -1  'True
      Top             =   210
      Width           =   1440
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "重庆市医疗保险"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   2190
      TabIndex        =   0
      Top             =   495
      Width           =   3465
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   1950
      Width           =   1020
   End
End
Attribute VB_Name = "frmIdentify重庆"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstr个人编号 As String
Private mint类别  As Long   '如果传入是表示0-门诊，1-住院；返回时表示11-普通门诊，13-特殊病门诊，14-急诊抢救，15-白内障摘除术，22-普通住院
Private mblnOK As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
            txtEdit(lngIndex).SetFocus
            Exit Sub
        End If
    Next
    
    If Trim(txtEdit(0).Text) = "" Then
        MsgBox "未输入个人帐户,不能通过验证！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstr个人编号 = Trim(txtEdit(0).Text)
    If mint类别 = 0 Then
        '门诊
        If opt类别(1).Value = True Then
            mint类别 = 13
        ElseIf opt类别(2).Value = True Then
            mint类别 = 14
        ElseIf opt类别(3).Value = True Then
            mint类别 = 15
        Else
            mint类别 = 11
        End If
    Else
        '住院
        mint类别 = 21
    End If
    
    mblnOK = True
    Unload Me
End Sub

Public Function GetIdentify(str个人编号 As String, int类别 As Integer) As Boolean
    mblnOK = False
    mstr个人编号 = str个人编号
    mint类别 = int类别
    
    If int类别 <> 0 Then
        '非门诊登记
        opt类别(0).Enabled = False
        opt类别(1).Enabled = False
        opt类别(2).Enabled = False
        opt类别(3).Enabled = False
    End If
    frmIdentify重庆.Show vbModal
    
    GetIdentify = mblnOK
    If mblnOK = True Then
        str个人编号 = mstr个人编号
        int类别 = mint类别
    End If
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub
