VERSION 5.00
Begin VB.Form frmTransfusionSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   Icon            =   "frmTransfusionSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6494.574
   ScaleMode       =   0  'User
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkTimeCall 
      Caption         =   "启用移动呼叫功能"
      Height          =   255
      Left            =   255
      TabIndex        =   30
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CheckBox chk接单穿刺 
      Caption         =   "接单后直接进入穿刺状态"
      Height          =   180
      Left            =   255
      TabIndex        =   4
      Top             =   1560
      Width           =   3045
   End
   Begin VB.CheckBox chkAutoReady 
      Caption         =   "通过查找功能找到病人后自动接单"
      Height          =   180
      Left            =   255
      TabIndex        =   3
      Top             =   1290
      Width           =   3045
   End
   Begin VB.Frame frmCardSet 
      Caption         =   "设备配置"
      Height          =   675
      Left            =   270
      TabIndex        =   28
      Top             =   2955
      Width           =   4470
      Begin VB.CommandButton cmdCardSet 
         Caption         =   "配置(&P)"
         Height          =   350
         Left            =   2985
         TabIndex        =   29
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.Frame fra 
      Caption         =   "请选择本工作站显示的单据类型"
      Height          =   660
      Left            =   270
      TabIndex        =   23
      Top             =   2250
      Width           =   4485
      Begin VB.CheckBox chkType 
         Caption         =   "治疗"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   27
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "输液"
         Height          =   195
         Index           =   1
         Left            =   1275
         TabIndex        =   26
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "注射"
         Height          =   195
         Index           =   2
         Left            =   2355
         TabIndex        =   25
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CheckBox chkType 
         Caption         =   "皮试"
         Height          =   195
         Index           =   3
         Left            =   3435
         TabIndex        =   24
         Top             =   315
         Value           =   1  'Checked
         Width           =   915
      End
   End
   Begin VB.TextBox txt皮试Time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3495
      MaxLength       =   4
      TabIndex        =   21
      ToolTipText     =   "最大提前时间60分钟"
      Top             =   960
      Width           =   465
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3495
      TabIndex        =   20
      Top             =   1140
      Width           =   465
   End
   Begin VB.TextBox txt输液Time 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3510
      MaxLength       =   4
      TabIndex        =   18
      ToolTipText     =   "最大提前时间60分钟"
      Top             =   690
      Width           =   465
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3510
      TabIndex        =   17
      Top             =   870
      Width           =   465
   End
   Begin VB.TextBox txt滴系数 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3675
      MaxLength       =   4
      TabIndex        =   15
      ToolTipText     =   "滴系数为10,15,20"
      Top             =   420
      Width           =   465
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3675
      TabIndex        =   14
      Top             =   600
      Width           =   465
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   3540
      TabIndex        =   13
      Top             =   315
      Width           =   465
   End
   Begin VB.TextBox txt滴速 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3540
      MaxLength       =   4
      TabIndex        =   11
      ToolTipText     =   "最大滴速100滴/分"
      Top             =   135
      Width           =   465
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3795
      Width           =   1100
   End
   Begin VB.CheckBox chkActLog 
      Caption         =   "允许其他人代行执行记录"
      Height          =   195
      Left            =   255
      TabIndex        =   0
      Top             =   135
      Width           =   2280
   End
   Begin VB.CheckBox chkFinish 
      Caption         =   "允许未收费病人完成执行"
      Height          =   195
      Left            =   255
      TabIndex        =   1
      Top             =   420
      Width           =   2280
   End
   Begin VB.CheckBox chk皮试 
      Caption         =   "填写皮试结果时验证身份"
      Height          =   195
      Left            =   255
      TabIndex        =   2
      Top             =   690
      Width           =   2280
   End
   Begin VB.TextBox txtRefresh 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   705
      MaxLength       =   4
      TabIndex        =   5
      ToolTipText     =   "最低刷新间隔为 30 秒，设置为 0 表示不自动刷新"
      Top             =   960
      Width           =   465
   End
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   705
      TabIndex        =   8
      Top             =   1140
      Width           =   465
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2490
      TabIndex        =   6
      Top             =   3795
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3630
      TabIndex        =   7
      Top             =   3795
      Width           =   1100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "皮试提前      分钟提醒"
      Height          =   180
      Left            =   2745
      TabIndex        =   22
      Top             =   960
      Width           =   1980
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输液提前      分钟提醒"
      Height          =   180
      Left            =   2760
      TabIndex        =   19
      Top             =   690
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认滴系数      "
      Height          =   180
      Left            =   2760
      TabIndex        =   16
      ToolTipText     =   "滴系数为10,15,20"
      Top             =   420
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "默认滴速      滴/分"
      Height          =   180
      Left            =   2790
      TabIndex        =   12
      ToolTipText     =   "最大滴速100滴/分"
      Top             =   135
      Width           =   1710
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "每      秒自动刷新清单"
      Height          =   180
      Left            =   480
      TabIndex        =   9
      ToolTipText     =   "最低刷新间隔为 30 秒，设置为 0 表示不自动刷新"
      Top             =   975
      Width           =   1980
   End
End
Attribute VB_Name = "frmTransfusionSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mstrPrivs As String
Public mlng科室ID As Long 'IN:当前执行科室ID
Public mblnOk As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCardSet_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOk_Click()
    Dim strPar As String, i As Long
    Dim strType As String
    Dim blnModify As Boolean
    '执行间范围
    blnModify = False
    If InStr(mstrPrivs, "参数设置") > 0 Then blnModify = True
    
    
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    Call zlDatabase.SetPara("医技刷新间隔", Val(txtRefresh.Text), glngSys, 1264, blnModify)
    
    '是否允许代行执行记录
    Call zlDatabase.SetPara("代行执行记录", chkActLog.Value, glngSys, 1264, blnModify)
    '是否允许完成未收费病人的项目
    Call zlDatabase.SetPara("未收费完成", chkFinish.Value, glngSys, 1264, blnModify)
    '填写皮试结果时验证身份
    Call zlDatabase.SetPara("皮试验证身份", chk皮试.Value, glngSys, 1264, blnModify)
    '接单后直接进入穿刺状态
    Call zlDatabase.SetPara("接单直接穿刺", chk接单穿刺.Value, glngSys, 1264)
    '移动呼叫
    Call zlDatabase.SetPara("移动呼叫", chkTimeCall.Value, glngSys, 1264)
    
    If Val(txt皮试Time.Text) < 0 Or Val(txt皮试Time.Text) > 60 Then txt皮试Time.Text = 0
    Call zlDatabase.SetPara("皮试提醒提前时间", Val(txt皮试Time.Text), glngSys, 1264, blnModify)
    
    
    If Val(txt滴速.Text) < 10 And Val(txt滴速.Text) > 100 Then txt滴速.Text = 40
    Call zlDatabase.SetPara("默认滴速", Val(txt滴速.Text), glngSys, 1264, blnModify)
    
    If InStr(",10,15,20,", "," & Val(txt滴系数.Text) & ",") <= 0 Then txt滴系数.Text = 20
    Call zlDatabase.SetPara("默认滴系数", Val(txt滴系数.Text), glngSys, 1264, blnModify)
    
    If Val(txt输液Time.Text) < 0 Or Val(txt输液Time.Text) > 60 Then txt输液Time.Text = 3
    Call zlDatabase.SetPara("输液提醒提前时间", Val(txt输液Time.Text), glngSys, 1264, blnModify)
    
    '2008-11-12
    strType = ""
    For i = 0 To chkType.Count - 1
        strType = strType & "," & chkType(i).Value
    Next
    Call zlDatabase.SetPara("显示单据种类", Mid(strType, 2), glngSys, 1264, blnModify)
    
    '2012-05-14 10.30 sp? 添加
    Call zlDatabase.SetPara("门诊输液自动接单", chkAutoReady.Value, glngSys, 1264, blnModify)
    
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then Call cmdHelp_Click
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim strType As String, i As Integer
    Dim intType As Integer '本机参数类型
    Dim blnModify As Boolean
    
    mblnOk = False
    blnModify = InStr(";" & mstrPrivs & ";", ";参数设置;") > 0
    
    cmdCardSet.Enabled = blnModify
    
    '修改:刘兴宏  修改了参数控制问题    日期:2008-06-12 10:58:11,主要加入了 Array(..),InStr(mstrPrivs, ";参数设置;")>0
    txtRefresh.Text = Val(zlDatabase.GetPara("医技刷新间隔", glngSys, 1264, "", Array(txtRefresh), blnModify))
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
      '是否允许代行执行记录
    chkActLog.Value = Val(zlDatabase.GetPara("代行执行记录", glngSys, 1264, "", Array(chkActLog), blnModify))
    
    '是否允许完成未收费病人的项目
    chkFinish.Value = Val(zlDatabase.GetPara("未收费完成", glngSys, 1264, "", Array(chkFinish), blnModify))
    
    '填写皮试结果时验证身份
    chk皮试.Value = Val(zlDatabase.GetPara("皮试验证身份", glngSys, 1264, "", Array(chk皮试), blnModify))
    
    '接单后直接进入穿刺状态
    chk接单穿刺.Value = Val(zlDatabase.GetPara("接单直接穿刺", glngSys, 1264, ""))
    
    '移动定时呼叫
    chkTimeCall.Value = Val(zlDatabase.GetPara("移动呼叫", glngSys, 1264))
        
    txt皮试Time.Text = Val(zlDatabase.GetPara("皮试提醒提前时间", glngSys, 1264, "", Array(txt皮试Time), blnModify))
    If Val(txt皮试Time.Text) < 0 Or Val(txt皮试Time.Text) > 60 Then txt皮试Time.Text = 0
    
    txt滴速.Text = Val(zlDatabase.GetPara("默认滴速", glngSys, 1264, "", Array(txt滴速), blnModify))
    If Val(txt滴速.Text) < 10 Or Val(txt滴速.Text) > 100 Then txt滴速.Text = 40
        
    txt滴系数.Text = Val(zlDatabase.GetPara("默认滴系数", glngSys, 1264, "", Array(txt滴系数), blnModify))
    If InStr(",10,15,20,", "," & Val(txt滴系数.Text) & ",") <= 0 Then txt滴系数.Text = 20

    txt输液Time.Text = Val(zlDatabase.GetPara("输液提醒提前时间", glngSys, 1264, "", Array(txt输液Time), blnModify))
    If Val(txt输液Time.Text) < 0 Or Val(txt输液Time.Text) > 60 Then txt输液Time.Text = 3
        
    '2008-11-12
    strType = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1", Array(Me.chkType(0), Me.chkType(1), Me.chkType(2), Me.chkType(3)), blnModify, intType)
    'strType = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
    For i = 0 To chkType.Count - 1
        chkType(i).Value = Val(Split(strType, ",")(i))
    Next
    '2012-05-14
    chkAutoReady.Value = Val(zlDatabase.GetPara("门诊输液自动接单", glngSys, 1264, "", Array(chkAutoReady), blnModify))
    
    '修改:刘兴宏  修改了参数控制问题    日期:2008-06-12 10:58:11,主要屏蔽以下代码
'    '权限设置的权限控制
'    If InStr(mstrPrivs, "参数设置") = 0 And intType = 15 Then
'        chkActLog.Enabled = False
'        chkFinish.Enabled = False
'        chk皮试.Enabled = False
'
'        txtRefresh.Enabled = False
'        txt滴速.Enabled = False
'        txt滴系数.Enabled = False
'        txt输液Time.Enabled = False
'        txt皮试Time.Enabled = False
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng科室ID = 0
    mstrPrivs = ""
End Sub

Private Sub txtRefresh_GotFocus()
    Call zlControl.TxtSelAll(txtRefresh)
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtRefresh_Validate(Cancel As Boolean)
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
End Sub

Private Sub txt滴速_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt滴速_Validate(Cancel As Boolean)
    If Val(txt滴速.Text) < 10 Or Val(txt滴速.Text) > 100 Then txt滴速.Text = 40
End Sub
Private Sub txt滴速_GotFocus()
    Call zlControl.TxtSelAll(txt滴速)
End Sub

Private Sub txt滴系数_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt滴系数_Validate(Cancel As Boolean)
    If InStr(",10,15,20,", "," & Val(txt滴系数.Text) & ",") <= 0 Then txt滴系数.Text = 20
End Sub
Private Sub txt滴系数_GotFocus()
    Call zlControl.TxtSelAll(txt滴系数)
End Sub

Private Sub txt输液Time_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt输液Time_Validate(Cancel As Boolean)
    If Val(txt输液Time.Text) < 0 Or Val(txt输液Time.Text) > 60 Then txt输液Time.Text = 3
End Sub
Private Sub txt输液Time_GotFocus()
    Call zlControl.TxtSelAll(txt输液Time)
End Sub

Private Sub txt皮试Time_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txt皮试Time_Validate(Cancel As Boolean)
    If Val(txt皮试Time.Text) < 0 Or Val(txt皮试Time.Text) > 60 Then txt皮试Time.Text = 0
End Sub
Private Sub txt皮试Time_GotFocus()
    Call zlControl.TxtSelAll(txt皮试Time)
End Sub
