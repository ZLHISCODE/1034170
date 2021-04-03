VERSION 5.00
Begin VB.Form frmOutStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "frmOutStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk缺省药房 
      Caption         =   "门诊医嘱下达强制缺省药房"
      Height          =   240
      Left            =   4350
      TabIndex        =   24
      Top             =   3180
      Width           =   2580
   End
   Begin VB.Frame fraEPR 
      Caption         =   "提醒设置"
      Height          =   1905
      Left            =   4350
      TabIndex        =   53
      Top             =   3495
      Width           =   4455
      Begin VB.CheckBox chkWarn 
         Caption         =   "输血反应"
         Height          =   195
         Index           =   5
         Left            =   3360
         TabIndex        =   33
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "用血审核"
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   32
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "备血完成"
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   31
         Top             =   1185
         Width           =   1035
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "处方审查"
         Height          =   195
         Index           =   2
         Left            =   2970
         TabIndex        =   30
         Top             =   885
         Width           =   1035
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "启用语音提示"
         Height          =   195
         Left            =   105
         TabIndex        =   34
         Top             =   1545
         Width           =   1470
      End
      Begin VB.CommandButton cmdSoundSet 
         Caption         =   "语音设置(&S)"
         Height          =   350
         Left            =   1650
         TabIndex        =   35
         Top             =   1455
         Width           =   1410
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "传染病"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   29
         Top             =   885
         Width           =   840
      End
      Begin VB.TextBox txtNotifyEPR 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   3
         TabIndex        =   26
         Text            =   "10"
         Top             =   330
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPR 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   585
         TabIndex        =   55
         Top             =   510
         Width           =   300
      End
      Begin VB.Frame fraNotifyEPRDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   585
         TabIndex        =   54
         Top             =   780
         Width           =   300
      End
      Begin VB.TextBox txtNotifyEPRDay 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   600
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "1"
         Top             =   600
         Width           =   300
      End
      Begin VB.CheckBox chkWarn 
         Caption         =   "危急值"
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   28
         Top             =   885
         Width           =   840
      End
      Begin VB.CheckBox chkNotifyEPR 
         Caption         =   "每    分钟自动刷新提醒区域中的内容"
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   345
         Width           =   3900
      End
      Begin VB.Label lblArea 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提醒内容:"
         Height          =   180
         Left            =   360
         TabIndex        =   57
         Top             =   885
         Width           =   810
      End
      Begin VB.Label lblNotifyEPRDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "将    天内完成的内容显示在提醒区域"
         Height          =   180
         Left            =   375
         TabIndex        =   56
         Top             =   615
         Width           =   3060
      End
   End
   Begin VB.CheckBox chkStaKB 
      Caption         =   "启用屏幕键盘"
      Height          =   255
      Left            =   330
      TabIndex        =   21
      Top             =   5070
      Width           =   1665
   End
   Begin VB.Frame fraBespeak 
      Caption         =   "预约挂号单打印方式"
      Height          =   2160
      Left            =   4350
      TabIndex        =   52
      Top             =   135
      Width           =   1920
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   315
         TabIndex        =   8
         Top             =   1575
         Width           =   1380
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   315
         TabIndex        =   6
         Top             =   450
         Width           =   900
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   7
         Top             =   1020
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.Frame fraReception 
      Caption         =   "病人接诊控制"
      Height          =   2145
      Left            =   6420
      TabIndex        =   49
      Top             =   150
      Width           =   2265
      Begin VB.OptionButton optMode 
         Caption         =   "不禁止"
         Height          =   240
         Index           =   0
         Left            =   645
         TabIndex        =   9
         Top             =   585
         Width           =   1005
      End
      Begin VB.OptionButton optMode 
         Caption         =   "禁止"
         Height          =   240
         Index           =   1
         Left            =   645
         TabIndex        =   10
         Top             =   900
         Width           =   855
      End
      Begin VB.OptionButton optMode 
         Caption         =   "提示"
         Height          =   240
         Index           =   2
         Left            =   645
         TabIndex        =   11
         Top             =   1230
         Width           =   750
      End
      Begin VB.TextBox txtReceptionTime 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   930
         TabIndex        =   12
         Text            =   "0"
         Top             =   1575
         Width           =   525
      End
      Begin VB.Label lblReceptionMode 
         Caption         =   "控制方式"
         Height          =   270
         Left            =   135
         TabIndex        =   51
         Top             =   330
         Width           =   825
      End
      Begin VB.Line line 
         X1              =   840
         X2              =   1545
         Y1              =   1785
         Y2              =   1785
      End
      Begin VB.Label lblReceptionTime 
         AutoSize        =   -1  'True
         Caption         =   "允许提前       分钟接诊"
         Height          =   180
         Left            =   120
         TabIndex        =   50
         Top             =   1590
         Width           =   2070
      End
   End
   Begin VB.OptionButton optAdd 
      Caption         =   "新增病历,切换到医嘱时新增医嘱"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   615
      TabIndex        =   18
      Top             =   4035
      Width           =   3015
   End
   Begin VB.CheckBox chk呼叫后接诊 
      Caption         =   "医生主动呼叫后才允许在队列中接诊"
      Height          =   195
      Left            =   330
      TabIndex        =   20
      Top             =   4785
      Value           =   1  'Checked
      Width           =   3360
   End
   Begin VB.CheckBox chk包含回诊病人 
      Caption         =   "医生呼叫人数限制含回诊病人"
      Height          =   180
      Left            =   4350
      TabIndex        =   22
      Top             =   2580
      Value           =   1  'Checked
      Width           =   2685
   End
   Begin VB.CheckBox chk挂号刷卡 
      Caption         =   "挂号必须刷卡提取病人"
      Height          =   255
      Left            =   330
      TabIndex        =   19
      Top             =   4410
      Width           =   2640
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   1605
      TabIndex        =   48
      Top             =   3045
      Width           =   465
   End
   Begin VB.TextBox txtQueuePatis 
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
      ForeColor       =   &H80000012&
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "3"
      ToolTipText     =   "表示本次医生最多能呼叫多少个病人来就诊,超过后，就不能再次呼叫;此参数需要配合分诊台模块的排队叫号模式为医生主动呼叫有效"
      Top             =   2880
      Width           =   465
   End
   Begin VB.CheckBox chkAutoAdd 
      Caption         =   "病人接诊后自动进行"
      Height          =   195
      Left            =   330
      TabIndex        =   16
      Top             =   3525
      Width           =   2640
   End
   Begin VB.CheckBox chk自动接诊 
      Caption         =   "查找到候诊病人之后自动接诊"
      Height          =   195
      Left            =   4350
      TabIndex        =   23
      Top             =   2880
      Width           =   2640
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   300
      TabIndex        =   36
      Top             =   5550
      Width           =   1500
   End
   Begin VB.CheckBox chkPrice 
      Caption         =   "允许挂号费用通过划价单收费"
      Height          =   195
      Left            =   330
      TabIndex        =   15
      Top             =   3180
      Value           =   1  'Checked
      Width           =   2640
   End
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   735
      TabIndex        =   46
      Top             =   2685
      Width           =   465
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
      ForeColor       =   &H80000012&
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   735
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "180"
      ToolTipText     =   "最低刷新间隔为 30 秒，设置为 0 表示不自动刷新"
      Top             =   2505
      Width           =   465
   End
   Begin VB.Frame Frame2 
      Caption         =   " 就诊参数 "
      Height          =   2190
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   4155
      Begin VB.CommandButton cmdYS 
         Caption         =   "…"
         Height          =   255
         Left            =   3645
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1755
         Width           =   255
      End
      Begin VB.TextBox txt接诊医生 
         Height          =   300
         Left            =   1020
         TabIndex        =   5
         Top             =   1725
         Width           =   2910
      End
      Begin VB.CheckBox chk分诊 
         Caption         =   "只接收已经分诊的病人"
         Height          =   195
         Left            =   1020
         TabIndex        =   4
         Top             =   1365
         Width           =   2100
      End
      Begin VB.ComboBox cbo科室 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   300
         Width           =   2910
      End
      Begin VB.CommandButton cmdSel 
         Caption         =   "…"
         Height          =   255
         Left            =   3645
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   690
         Width           =   255
      End
      Begin VB.ComboBox cbo范围 
         ForeColor       =   &H80000012&
         Height          =   300
         ItemData        =   "frmOutStationSetup.frx":000C
         Left            =   1020
         List            =   "frmOutStationSetup.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "接诊的病人范围"
         Top             =   1005
         Width           =   2910
      End
      Begin VB.TextBox txt诊室 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1020
         MaxLength       =   20
         TabIndex        =   2
         Top             =   660
         Width           =   2910
      End
      Begin VB.Label lblEditDept 
         AutoSize        =   -1  'True
         Caption         =   "接诊科室"
         Height          =   180
         Left            =   255
         TabIndex        =   0
         Top             =   360
         Width           =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   45
         X2              =   4090
         Y1              =   1650
         Y2              =   1650
      End
      Begin VB.Label lbl医生 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接诊医生"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   1785
         Width           =   720
      End
      Begin VB.Label lbl诊室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医生诊室"
         Height          =   180
         Left            =   255
         TabIndex        =   39
         Top             =   705
         Width           =   720
      End
      Begin VB.Label lbl范围 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "接诊范围"
         Height          =   180
         Left            =   225
         TabIndex        =   41
         Top             =   1065
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   45
         X2              =   4090
         Y1              =   1635
         Y2              =   1635
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   38
      Top             =   5550
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   37
      Top             =   5550
      Width           =   1100
   End
   Begin VB.OptionButton optAdd 
      Caption         =   "新增医嘱,切换到病历时新增病历"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   615
      TabIndex        =   17
      Top             =   3765
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.Label lblQueuePatis 
      AutoSize        =   -1  'True
      Caption         =   "医生最多能呼叫      人"
      Height          =   180
      Left            =   330
      TabIndex        =   47
      ToolTipText     =   "表示本次医生最多能呼叫多少个病人来就诊,超过后，就不能再次呼叫;此参数需要配合分诊台模块的排队叫号模式为医生主动呼叫有效"
      Top             =   2880
      Width           =   1980
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      X1              =   -15
      X2              =   9000
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6780
      Y1              =   5415
      Y2              =   5415
   End
   Begin VB.Label lblRefresh 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "每隔      秒自动刷新候诊/转诊病人清单"
      Height          =   180
      Left            =   345
      TabIndex        =   44
      ToolTipText     =   "最低刷新间隔为 30 秒，设置为 0 表示不自动刷新"
      Top             =   2520
      Width           =   3330
   End
End
Attribute VB_Name = "frmOutStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String
Private mstrLike As String
Private mbln挂号按排 As Boolean '根据参数：挂号排班模式  确定诊室选择范围，true新版，false老版

Private Enum Enum_chkWarn
    chkD危急值 = 0
    chkD传染病 = 1
    chkD处方审查 = 2
    chkD备血完成 = 3
    chkD用血审核 = 4
    chkD输血反应 = 5
End Enum

Private Sub cbo范围_Click()
    '本人号或本科室时
    chk分诊.Visible = cbo范围.ListIndex = 0 Or cbo范围.ListIndex = 2
End Sub


Private Sub chkAutoAdd_Click()
    If chkAutoAdd.Value = 1 Then
        optAdd(0).Enabled = True
        optAdd(1).Enabled = True
    Else
        optAdd(0).Enabled = False
        optAdd(1).Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
End Sub

Private Sub cmdOK_Click()
    Dim str病人接诊控制 As String '问题号:57566
    Dim blnHavePara As Boolean  '是否有参数设置权限
    Dim i As Integer
    Dim strTmp As String
    
    If txt诊室.Text = "" Then
        MsgBox "请设置医生的诊室。", vbInformation, gstrSysName
        txt诊室.SetFocus: Exit Sub
    End If
    If txt接诊医生.Text = "" Then
        MsgBox "请接诊医生。", vbInformation, gstrSysName
        txt接诊医生.SetFocus: Exit Sub
    End If
    If cbo科室.ListIndex < 0 Then
        MsgBox "接诊科室必须选择,请检查", vbInformation + vbOKOnly, gstrSysName
        cbo科室.SetFocus
        Exit Sub
    End If
    blnHavePara = InStr(1, ";" & mstrPrivs & ";", ";参数设置;") > 0
    
    Call zlDatabase.SetPara("本地诊室", Me.txt诊室.Text, glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊范围", Me.cbo范围.ItemData(Me.cbo范围.ListIndex), glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("接诊医生", Me.txt接诊医生.Text, glngSys, p门诊医生站, blnHavePara)
    '问题:38603
    Call zlDatabase.SetPara("挂号必须刷卡", chk挂号刷卡.Value, glngSys, p门诊医生站, blnHavePara)
    
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫对象=2时有效
    If txtQueuePatis.Enabled Then
        Call zlDatabase.SetPara("医生就诊人数", Val(Me.txtQueuePatis.Text), glngSys, p门诊医生站, blnHavePara)
    End If
    '接诊科室
    Call zlDatabase.SetPara("接诊科室", cbo科室.ItemData(cbo科室.ListIndex), glngSys, p门诊医生站, blnHavePara)
    
    '只接收已经分诊的病人
    Call zlDatabase.SetPara("只接收已经分诊的病人", chk分诊.Value, glngSys, p门诊医生站, blnHavePara)

    '候诊病人刷新间隔
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    Call zlDatabase.SetPara("候诊刷新间隔", Val(txtRefresh.Text), glngSys, p门诊医生站, blnHavePara)
    
    '挂号费用不通过划价单收费
    Call zlDatabase.SetPara("允许挂号划价单", chkPrice.Value, glngSys, p门诊医生站, blnHavePara)
    
    '找到病人后自动接诊
    Call zlDatabase.SetPara("找到病人后自动接诊", chk自动接诊.Value, glngSys, p门诊医生站, blnHavePara)
    '接诊后自动进行
    If optAdd(1).Value And optAdd(1).Enabled Then
        Call zlDatabase.SetPara("接诊后自动进行", 2, glngSys, p门诊医生站, blnHavePara)
    Else
        Call zlDatabase.SetPara("接诊后自动进行", chkAutoAdd.Value, glngSys, p门诊医生站, blnHavePara)
    End If
    '问题:44250
    Call zlDatabase.SetPara("就诊人数含回诊", chk包含回诊病人.Value, glngSys, p门诊医生站, blnHavePara)
    '医生主动呼叫后才允许接诊
    Call zlDatabase.SetPara("医生主动呼叫后才允许接诊", chk呼叫后接诊.Value, glngSys, p门诊医生站, blnHavePara)
    '启用屏幕键盘
    Call zlDatabase.SetPara("启用屏幕键盘", chkStaKB.Value, glngSys, p门诊医生站, blnHavePara)
    '问题号:57566
    If optMode(0).Value = True Then
        str病人接诊控制 = "0|0"
    ElseIf optMode(1).Value = True Then
        str病人接诊控制 = "1|" & Nvl(txtReceptionTime.Text, "0")
    ElseIf optMode(2).Value = True Then
        str病人接诊控制 = "2|" & Nvl(txtReceptionTime.Text, "0")
    End If
    zlDatabase.SetPara "病人接诊控制", str病人接诊控制, glngSys, p门诊医生站, blnHavePara
    
    '56274
    For i = 0 To optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "预约挂号单打印方式", i, glngSys, p门诊医生站, blnHavePara
            Exit For
        End If
    Next
    
    Call zlDatabase.SetPara("门诊医嘱下达强制缺省药房", chk缺省药房.Value, glngSys, p门诊医嘱下达, blnHavePara)
    
    Call zlDatabase.SetPara("自动刷新病历审阅间隔", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("自动刷新病历审阅天数", Val(txtNotifyEPRDay.Text), glngSys, p门诊医生站, blnHavePara)
    strTmp = ""
    For i = chkD危急值 To chkD输血反应
        strTmp = strTmp & chkWarn(i).Value
    Next
    Call zlDatabase.SetPara("自动刷新内容", strTmp, glngSys, p门诊医生站, blnHavePara)
    Call zlDatabase.SetPara("启用语音提示", chkSound.Value, glngSys, p门诊医生站, blnHavePara)
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSel_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If txt诊室.Tag <> txt诊室 Then Exit Sub '由txt诊室的Validate事件处理
    
    If mbln挂号按排 Then
        strSQL = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id" & vbNewLine & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    Else
        strSQL = "Select Distinct e.编码 As ID,e.名称,e.简码" & vbNewLine & _
               "From 门诊诊室 E, 挂号安排诊室 D, 挂号安排 C, 部门人员 A, 上机人员表 B" & vbNewLine & _
               "Where a.人员id = b.人员id And b.用户名 = User And c.科室id = a.部门id And c.Id = d.号表id And e.名称 = d.门诊诊室 " & _
               " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null)"
    End If
    '如果没有查找到数据，则读取出所有的门诊诊室供选择
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    End If

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "门诊诊室", , , , , , , txt诊室.Left, txt诊室.Top, txt诊室.Height, , , True)
    If Not rsTmp Is Nothing Then
        txt诊室.Tag = rsTmp("名称"): txt诊室 = txt诊室.Tag
        If cbo范围.Enabled And cbo范围.Visible Then cbo范围.SetFocus
    End If
End Sub

Private Sub cmdSoundSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 0)
End Sub

Private Sub cmdYS_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnCanle As Boolean
    If txt接诊医生.Tag <> txt接诊医生 Then Exit Sub '由txt医生的Validate事件处理
            
    strSQL = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床'" & _
        " And B.部门ID In (Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.简码"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "接诊医生", False, "", "", False, False, False, 0, 0, txt接诊医生.Height, blnCanle, False, True, UserInfo.ID)
    If blnCanle Then Exit Sub
    If Not rsTmp Is Nothing Then txt接诊医生.Tag = rsTmp("名称"): txt接诊医生 = txt接诊医生.Tag: Me.cmdOK.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String
    Dim blnSetup As Boolean
    Dim i As Long
    Dim str病人接诊控制 As String  '问题号:57566
    Dim intType As Integer
    Dim strNotify As String
    Dim str诊室 As String
    
    blnSetup = InStr(1, ";" & mstrPrivs & ";", ";参数设置;") > 0
    gblnOK = False
    mstrLike = IIf(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "") '输入匹配方式
    
    On Error Resume Next
    str诊室 = zlDatabase.GetPara("本地诊室", glngSys, p门诊医生站, "", Array(lbl诊室, txt诊室, cmdSel), blnSetup)
    On Error GoTo 0
    
    On Error GoTo errH
    '读取病人缺省科室范围
    strPar = zlDatabase.GetPara("接诊科室", glngSys, p门诊医生站, "", Array(lblEditDept, cbo科室), blnSetup)
    
    strSQL = "Select Distinct B.ID,B.编码,B.名称,A.缺省" & _
        " From 部门人员 A,部门表 B,部门性质说明 C" & _
        " Where A.部门ID=B.ID And B.ID=C.部门ID And C.服务对象 In(1,3) And C.工作性质='临床'" & _
        " And (B.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or B.撤档时间 is Null)" & _
        " And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null) And A.人员ID=[1]" & _
        " Order by B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    For i = 1 To rsTmp.RecordCount
        cbo科室.AddItem rsTmp!名称
        cbo科室.ItemData(cbo科室.NewIndex) = rsTmp!ID
        If rsTmp!ID = Val(strPar) Then
            cbo科室.ListIndex = cbo科室.NewIndex
        ElseIf Nvl(rsTmp!缺省, 0) = 1 And cbo科室.ListIndex = -1 Then
            cbo科室.ListIndex = cbo科室.NewIndex
        End If
        rsTmp.MoveNext
    Next
    Me.cbo范围.ListIndex = Val(zlDatabase.GetPara("接诊范围", glngSys, p门诊医生站, "2", Array(lbl范围, cbo范围), blnSetup)) - 1
    
    strSQL = "Select 1 From 门诊诊室 E where e.名称=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str诊室)
    If Not rsTmp.EOF Then
        txt诊室.Text = str诊室
        txt诊室.Tag = str诊室
    End If
    
    '可以选择对其它医生就诊的病人进行就诊
    If InStr(mstrPrivs, "续诊病人") > 0 Then
        '可以选择本科室下的医生
        cmdYS.Enabled = True
        txt接诊医生.Enabled = True
    Else
        cmdYS.Enabled = False
        txt接诊医生.Enabled = False
    End If
    txt接诊医生.Tag = zlDatabase.GetPara("接诊医生", glngSys, p门诊医生站, UserInfo.姓名, Array(lbl医生, txt接诊医生, cmdYS), blnSetup)
    txt接诊医生.Text = txt接诊医生.Tag
    
    '问题:38603
    chk挂号刷卡.Value = IIf(Val(zlDatabase.GetPara("挂号必须刷卡", glngSys, p门诊医生站, "0", Array(chk挂号刷卡), blnSetup)) = 1, 1, 0)
    '刘兴洪:应用于排队叫号的呼叫人次:需要配合分诊台模块的排队叫号模式为１并且有排队呼叫站点=1时有效
    txtQueuePatis.Text = Val(zlDatabase.GetPara("医生就诊人数", glngSys, p门诊医生站, 3, Array(lblQueuePatis, txtQueuePatis), blnSetup))
    If txtQueuePatis.Enabled Then
        txtQueuePatis.Enabled = CheckDoctorPatisIsValid
    End If
    
    '只接收已经分诊的病人
    chk分诊.Value = Val(zlDatabase.GetPara("只接收已经分诊的病人", glngSys, p门诊医生站, , Array(chk分诊), blnSetup))
    
    '候诊病人刷新间隔
    txtRefresh.Text = Val(zlDatabase.GetPara("候诊刷新间隔", glngSys, p门诊医生站, 180, Array(lblRefresh, txtRefresh), blnSetup))
    If Val(txtRefresh.Text) <> 0 And Val(txtRefresh.Text) < 30 Then txtRefresh.Text = 30
    
    '挂号费用不通过划价单收费
    chkPrice.Value = Val(zlDatabase.GetPara("允许挂号划价单", glngSys, p门诊医生站, 1, Array(chkPrice), blnSetup))
    
    '找到病人后自动接诊
    chk自动接诊.Value = Val(zlDatabase.GetPara("找到病人后自动接诊", glngSys, p门诊医生站, , Array(chk自动接诊), blnSetup))
    
    '接诊后自动进行
    strPar = Val(zlDatabase.GetPara("接诊后自动进行", glngSys, p门诊医生站, , Array(chkAutoAdd, optAdd(0), optAdd(1)), blnSetup))
    If strPar = 2 Then
        chkAutoAdd.Value = 1
        optAdd(1).Value = True
    Else
        chkAutoAdd.Value = strPar
    End If
    '问题:44250
    chk包含回诊病人.Value = Val(zlDatabase.GetPara("就诊人数含回诊", glngSys, p门诊医生站, 1, Array(chk包含回诊病人), blnSetup))
    '医生主动呼叫后才允许接诊
    chk呼叫后接诊.Value = Val(zlDatabase.GetPara("医生主动呼叫后才允许接诊", glngSys, p门诊医生站, 1, Array(chk呼叫后接诊), blnSetup))
    '启用屏幕键盘
    chkStaKB.Value = Val(zlDatabase.GetPara("启用屏幕键盘", glngSys, p门诊医生站, , Array(chkStaKB), blnSetup))
    
    '问题号:57566
    '病人接诊控制
    str病人接诊控制 = zlDatabase.GetPara("病人接诊控制", glngSys, p门诊医生站, , Array(optMode(0), optMode(1), optMode(2), txtReceptionTime, lblReceptionMode, lblReceptionTime), blnSetup)
    If str病人接诊控制 <> "" Then
        If Split(str病人接诊控制, "|")(0) = "0" Then
            optMode(0).Value = True
        ElseIf Split(str病人接诊控制, "|")(0) = "1" Then
            optMode(1).Value = True
        ElseIf Split(str病人接诊控制, "|")(0) = "2" Then
            optMode(2).Value = True
        End If
        txtReceptionTime.Text = Split(str病人接诊控制 & "|", "|")(1)
    End If
    '问题:56274
    i = Val(zlDatabase.GetPara("预约挂号单打印方式", glngSys, p门诊医生站, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2)), blnSetup))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
    
    '消息提醒刷新
    strPar = zlDatabase.GetPara("自动刷新病历审阅间隔", glngSys, p门诊医生站, , Array(chkNotifyEPR), blnSetup, intType)
    If Val(strPar) > 0 Then
        chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
    End If
 
    If (intType = 3 Or intType = 15) And Not blnSetup Then
        txtNotifyEPR.Enabled = False
    End If
    
    strPar = zlDatabase.GetPara("自动刷新病历审阅天数", glngSys, p门诊医生站, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), blnSetup)
    txtNotifyEPRDay.Text = Val(strPar)
        
    strNotify = zlDatabase.GetPara("自动刷新内容", glngSys, p门诊医生站, , Array(chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), lblArea), blnSetup)
    chkWarn(chkD危急值).Value = Val(Mid(strNotify, 1, 1))
    chkWarn(chkD传染病).Value = Val(Mid(strNotify, 2, 1))
    chkWarn(chkD处方审查).Value = Val(Mid(strNotify, 3, 1))
    chkWarn(chkD备血完成).Value = Val(Mid(strNotify, 4, 1))
    chkWarn(chkD备血完成).Visible = gbln血库系统
    chkWarn(chkD用血审核).Value = Val(Mid(strNotify, 5, 1))
    chkWarn(chkD用血审核).Visible = gbln血库系统
    chkWarn(chkD输血反应).Value = Val(Mid(strNotify, 6, 1))
    chkWarn(chkD输血反应).Visible = gbln血库系统
    If InStr(mstrPrivs, "参数设置") = 0 Then
        chkWarn(chkD危急值).Enabled = False
        chkWarn(chkD传染病).Enabled = False
        chkWarn(chkD处方审查).Enabled = False
        chkWarn(chkD备血完成).Enabled = False
        chkWarn(chkD用血审核).Enabled = False
        chkWarn(chkD输血反应).Enabled = False
    End If
    chkSound.Value = Val(zlDatabase.GetPara("启用语音提示", glngSys, p门诊医生站, , Array(chkSound, cmdSoundSet), blnSetup))

    '门诊医嘱下达强制缺省药房
    chk缺省药房.Value = Val(zlDatabase.GetPara("门诊医嘱下达强制缺省药房", glngSys, p门诊医嘱下达, "1", Array(chk缺省药房), blnSetup))

    strPar = ""
    mbln挂号按排 = False
    strPar = zlDatabase.GetPara(256, glngSys) & "|"
    If 0 <> Val(Split(strPar, "|")(0)) Then
        If Split(strPar, "|")(1) <> "" Then
            strPar = Format(Split(strPar, "|")(1), "YYYY-MM-DD")
            If Format(zlDatabase.Currentdate, "YYYY-MM-DD") >= strPar Then
                mbln挂号按排 = True
            End If
        End If
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optMode_Click(Index As Integer)
    '问题号:57566
    Select Case Index
        Case 0
            txtReceptionTime.Text = 0: txtReceptionTime.Enabled = False
        Case Else
            txtReceptionTime.Enabled = True And InStr(mstrPrivs, ";参数设置;") > 0
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

Private Sub txt接诊医生_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean

    If txt接诊医生.Tag = txt接诊医生 Then Exit Sub

    strSQL = "Select Distinct A.编号 as ID,A.姓名 as 名称,A.简码" & _
        " From 人员表 A,部门人员 B,人员性质说明 C,部门性质说明 D" & _
        " Where A.ID=B.人员ID And A.ID=C.人员ID And B.部门ID=D.部门ID" & _
        " And C.人员性质||''='医生' And D.服务对象 IN(1,3) And D.工作性质||''='临床'" & _
        " And B.部门ID In(Select 部门ID From 部门人员 Where 人员ID=[1])" & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (Upper(A.编号) Like [2] Or Upper(A.简码) Like [3] Or Upper(A.姓名) Like [3])" & _
        " Order by A.简码"
        
    vRect = GetControlRect(txt接诊医生.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "接诊医生", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt接诊医生.Height, blnCancel, False, True, UserInfo.ID, UCase(txt接诊医生.Text) & "%", mstrLike & UCase(txt接诊医生.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt接诊医生.Tag = rsTmp("名称")
        txt接诊医生 = txt接诊医生.Tag
    Else
        txt接诊医生.Tag = ""
        txt接诊医生 = ""
        Cancel = blnCancel
    End If
End Sub

Private Sub txt接诊医生_GotFocus()
    Call zlControl.TxtSelAll(txt接诊医生)
End Sub

Private Sub txt接诊医生_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt接诊医生 = "" Then txt接诊医生.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_GotFocus()
    Call zlControl.TxtSelAll(txt诊室)
End Sub

Private Sub txt诊室_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txt诊室 = "" Then txt诊室.Tag = "1"
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt诊室_Validate(Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, blnCancel As Boolean
    
    If txt诊室.Tag = txt诊室 Then Exit Sub
    
    If mbln挂号按排 Then
        strSQL = "Select Distinct a.Id, a.名称, a.简码" & vbNewLine & _
            " From 门诊诊室 A, 门诊诊室适用科室 B, 部门人员 C, 上机人员表 D" & vbNewLine & _
            " Where a.Id = b.诊室id And b.科室id = c.部门id And c.人员id = d.人员id" & vbNewLine & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])" & _
            "       And d.用户名 = User And (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)"
    Else
        strSQL = "Select Distinct e.编码 As ID,e.名称,e.简码" & vbNewLine & _
                "From 门诊诊室 E, 挂号安排诊室 D, 挂号安排 C, 部门人员 A, 上机人员表 B" & vbNewLine & _
                "Where a.人员id = b.人员id And b.用户名 = User And c.科室id = a.部门id And c.Id = d.号表id And e.名称 = d.门诊诊室 " & _
                " And (Upper(E.编码) Like [1] Or Upper(E.简码) Like [2] Or Upper(E.名称) Like [2])" & _
                " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) "
    End If
        
    '如果没有查找到数据，则读取出所有的门诊诊室供选择
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    If rsTmp.EOF Then
        strSQL = "Select a.Id, a.名称, a.简码 From 门诊诊室 A Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 Is Null)" & _
            " And (Upper(a.编码) Like [1] Or Upper(a.简码) Like [2] Or Upper(a.名称) Like [2])"
    End If

    vRect = GetControlRect(txt诊室.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "门诊诊室", False, "", "", False, False, True, _
        vRect.Left, vRect.Top, txt诊室.Height, blnCancel, False, True, UCase(txt诊室.Text) & "%", mstrLike & UCase(txt诊室.Text) & "%")
    If Not rsTmp Is Nothing Then
        txt诊室.Tag = rsTmp("名称")
        txt诊室 = txt诊室.Tag
    Else
        txt诊室.Tag = ""
        txt诊室 = ""
        Cancel = blnCancel
    End If
End Sub
Private Sub txtReceptionTime_GotFocus()
    zlControl.TxtSelAll txtReceptionTime
End Sub
Private Sub txtReceptionTime_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtReceptionTime_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtReceptionTime, KeyAscii, m数字式
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
