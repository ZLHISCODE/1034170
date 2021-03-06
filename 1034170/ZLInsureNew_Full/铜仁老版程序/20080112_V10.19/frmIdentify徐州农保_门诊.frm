VERSION 5.00
Begin VB.Form frmIdentify徐州农保_门诊 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3285
   Icon            =   "frmIdentify徐州农保_门诊.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   810
      TabIndex        =   2
      Top             =   1080
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   1995
      TabIndex        =   3
      Top             =   1080
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   105
      Left            =   -60
      TabIndex        =   4
      Top             =   810
      Width           =   3555
   End
   Begin VB.TextBox txt姓名 
      Height          =   300
      Left            =   780
      TabIndex        =   1
      Top             =   330
      Width           =   2175
   End
   Begin VB.Label lbl姓名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   390
      Width           =   360
   End
End
Attribute VB_Name = "frmIdentify徐州农保_门诊"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytType As Byte, mstrOther As String, mstrPatient As String

Public Function GetPatient(bytType As Byte) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mstrPatient = ""
    mstrOther = ""
    
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngSequence As Long
    
    If Trim(txt姓名.Text) = "" Then
        MsgBox "请输入病人姓名！", vbInformation, gstrSysName
        txt姓名.SetFocus
        Exit Sub
    End If
    
    '因病人信息还未产生，只有以人员表的序列做为医保号与卡号
    lngSequence = zlDatabase.GetNextId("人员表")
    
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = lngSequence & ";"                                     '0 卡号
    mstrPatient = mstrPatient & lngSequence & ";"                       '1 医保帐号
    mstrPatient = mstrPatient & ";"                                     '2 密码
    mstrPatient = mstrPatient & Me.txt姓名.Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & ";"                                     '4 性别
    mstrPatient = mstrPatient & ";"                                     '5 出生日期
    mstrPatient = mstrPatient & ";"                                     '6 身份证
    mstrPatient = mstrPatient & ";"                                     '7 单位名称/编码
    
    mstrOther = mstrOther & ";"                                         '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & ";"                                         '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & ";"
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & ";"                                         '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & ";"                                         '20 进入统筹累计
    mstrOther = mstrOther & ";"                                         '21 统筹报销累计
    mstrOther = mstrOther & ";"                                         '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & ";"                                         '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    Me.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        Call SendKeys("{Tab}")
    End If
End Sub

