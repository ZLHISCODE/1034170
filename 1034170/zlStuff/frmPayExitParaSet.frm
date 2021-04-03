VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPayExitParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8100
   Icon            =   "frmPayExitParaSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   8100
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Caption         =   "  其他 "
      Height          =   1272
      Index           =   2
      Left            =   5160
      TabIndex        =   18
      Top             =   960
      Width           =   2850
      Begin VB.CheckBox chk发生时间 
         Caption         =   "卫材医嘱按发生时间过滤"
         Height          =   345
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   2340
      End
      Begin VB.CheckBox chkSign 
         Caption         =   "领料人签名"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   1485
      End
      Begin VB.CheckBox chk销帐 
         Caption         =   "退料时自动将记帐费用销帐"
         Height          =   360
         Left            =   240
         TabIndex        =   8
         Top             =   255
         Width           =   2520
      End
   End
   Begin VB.Frame fra设备定义 
      Caption         =   "  智能卡及其他设备定义 "
      Height          =   735
      Left            =   75
      TabIndex        =   23
      Top             =   5520
      Width           =   7950
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   240
         TabIndex        =   24
         Top             =   250
         Width           =   1500
      End
   End
   Begin VB.Frame fra 
      Caption         =   "  发料控制 "
      Height          =   1545
      Index           =   3
      Left            =   75
      TabIndex        =   19
      Top             =   2400
      Width           =   7950
      Begin VB.CheckBox chk汇总发料 
         Caption         =   "发料时汇总销帐申请记录"
         Height          =   180
         Left            =   2400
         TabIndex        =   36
         Top             =   275
         Width           =   2655
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "营养"
         Enabled         =   0   'False
         Height          =   180
         Index           =   6
         Left            =   2280
         TabIndex        =   31
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "治疗"
         Enabled         =   0   'False
         Height          =   180
         Index           =   5
         Left            =   1440
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "手术"
         Enabled         =   0   'False
         Height          =   180
         Index           =   4
         Left            =   600
         TabIndex        =   29
         Top             =   1200
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "检验"
         Enabled         =   0   'False
         Height          =   180
         Index           =   3
         Left            =   3240
         TabIndex        =   28
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "检查"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   27
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "护理"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   1440
         TabIndex        =   26
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkDeptType 
         Caption         =   "临床"
         Enabled         =   0   'False
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   25
         Top             =   900
         Width           =   735
      End
      Begin VB.CheckBox chkSendByNo 
         Caption         =   "按单据号发料"
         Height          =   180
         Left            =   5160
         TabIndex        =   22
         Top             =   240
         Width           =   2130
      End
      Begin VB.CheckBox chk病区 
         Caption         =   "按病区发料时包含非病人科室开单的记录"
         Height          =   420
         Left            =   240
         TabIndex        =   21
         Top             =   500
         Width           =   4650
      End
      Begin VB.CheckBox Chk是否自动缺料检查 
         Caption         =   "是否自动缺料检查"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   275
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  业务类型 "
      Height          =   1272
      Left            =   75
      TabIndex        =   17
      Top             =   960
      Width           =   3090
      Begin VB.ComboBox cbo收费处方 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   840
         Width           =   2280
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "收费单(&S)"
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   0
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "记帐单(&J)"
         Height          =   285
         Index           =   1
         Left            =   1850
         TabIndex        =   1
         Top             =   240
         Width           =   1150
      End
      Begin VB.CheckBox chk业务 
         Caption         =   "记帐表(&B)"
         Height          =   285
         Index           =   2
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   1150
      End
      Begin VB.Label lbl收费处方 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "收费处方"
         Height          =   420
         Left            =   120
         TabIndex        =   34
         Top             =   825
         Width           =   465
      End
      Begin VB.Label lbl单据类型 
         Caption         =   "单据类型"
         Height          =   420
         Left            =   120
         TabIndex        =   32
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   6480
      Width           =   8775
   End
   Begin VB.Frame fraLine 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   705
      Width           =   8775
   End
   Begin VB.Frame fra 
      Caption         =   "  打印及票据设置 "
      Height          =   1305
      Index           =   1
      Left            =   75
      TabIndex        =   14
      Top             =   4080
      Width           =   7950
      Begin VB.ComboBox cbo发料后 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   360
         Width           =   2280
      End
      Begin VB.ComboBox cbo退料后 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   360
         Width           =   2280
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "票据打印设置"
         Height          =   360
         Left            =   3360
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   750
         Width           =   1875
      End
      Begin VB.ComboBox cbo票据设置 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   780
         Width           =   2280
      End
      Begin VB.Label lbl发料单 
         AutoSize        =   -1  'True
         Caption         =   "发料单"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lbl退料单 
         AutoSize        =   -1  'True
         Caption         =   "退料单"
         Height          =   180
         Left            =   3360
         TabIndex        =   40
         Top             =   420
         Width           =   540
      End
      Begin VB.Label lbl票据 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "票据(&S)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "  缺省单位 "
      Height          =   1272
      Index           =   0
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   1770
      Begin VB.CheckBox chk单位 
         Caption         =   "包装单位(&2)"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   660
         Width           =   1452
      End
      Begin VB.CheckBox chk单位 
         Caption         =   "散装单位(&1)"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   348
         Width           =   1452
      End
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   11
      Top             =   6735
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   10
      Top             =   6735
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5385
      TabIndex        =   9
      Top             =   6735
      Width           =   1100
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   7344
      Top             =   -48
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":030A
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":08A4
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":0E3E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayExitParaSet.frx":13D8
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmPayExitParaSet.frx":1972
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "根据下面选项目,设置相关的打印、发料单位和相关票据的设置。"
      Height          =   390
      Index           =   0
      Left            =   735
      TabIndex        =   12
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmPayExitParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private mblnExit As Boolean
Private mlngModule As Long
Private mstrPrivs As String
Private mblnHavePriv As Boolean
Private Const mstrAllType As String = "临床,护理,检查,检验,手术,治疗,营养"


Private Sub cbo票据设置_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkDeptType_Click(Index As Integer)
    Dim n As Integer
    Dim blnAllUnselect As Boolean
    
    '至少要选择一个
    blnAllUnselect = True
    For n = 0 To chkDeptType.Count - 1
        If chkDeptType(n).Value = 1 Then
            blnAllUnselect = False
            Exit For
        End If
    Next
    If blnAllUnselect = True Then
        chkDeptType(Index).Value = 1
    End If
End Sub

Private Sub chk病区_Click()
    Dim n As Integer

    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = (chk病区.Value = 1)
        If chk病区.Tag = "0" Then
            chkDeptType(n).Value = 1
        End If
    Next
End Sub



Private Sub chk打印_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk单位_Click(Index As Integer)
    
    If chk单位(Index).Value = 1 Then
        chk单位(IIf(Index = 1, 0, 1)).Value = 0
    Else
        chk单位(IIf(Index = 1, 0, 1)).Value = 1
    End If
End Sub

Private Sub chk单位_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab

End Sub



 
 
Private Sub chk销帐_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk业务_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub CmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1723)
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub
Private Function SaveSet() As Boolean
    '------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴宏
    '日期:2007/12/24
    '------------------------------------------------------------------------------------------
    Dim str业务类型 As String
    Dim str病区发料 As String
    Dim n As Integer
    
    str业务类型 = IIf(chk业务(0).Value = 1, "24", "0")
    str业务类型 = str业务类型 & IIf(chk业务(1).Value = 1, ",25", ",0")
    str业务类型 = str业务类型 & IIf(chk业务(2).Value = 1, ",26", ",0")
    
    '病区发药
    If chk病区.Value = 0 Then
        str病区发料 = ""
    Else
        For n = 0 To chkDeptType.Count - 1
            If chkDeptType(n).Value = 0 Then
                str病区发料 = IIf(str病区发料 = "", "", str病区发料 & ",") & chkDeptType(n).Caption
            End If
        Next
        If str病区发料 = "" Then
            str病区发料 = mstrAllType
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    gcnOracle.BeginTrans
   
    Call zlDatabase.SetPara("发料打印提醒方式", cbo发料后.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("退料打印提醒方式", cbo退料后.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("查询业务类型", str业务类型, glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材单位", IIf(chk单位(1).Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("自动销帐", IIf(chk销帐.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("缺料检查", IIf(Chk是否自动缺料检查.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("病区发料方式", str病区发料, glngSys, mlngModule)
    Call zlDatabase.SetPara("按单据号发料", chkSendByNo.Value, glngSys, mlngModule)
    Call zlDatabase.SetPara("收费处方显示方式", cbo收费处方.ListIndex, glngSys, mlngModule)
    Call zlDatabase.SetPara("发料时汇总退料销帐记录", chk汇总发料.Value, glngSys, mlngModule)
    '59655
    Call zlDatabase.SetPara("领料人签名", IIf(chkSign.Value = 1, 1, 0), glngSys, mlngModule)
    Call zlDatabase.SetPara("卫材医嘱按发生时间过滤", IIf(chk发生时间.Value = 1, 1, 0), glngSys, mlngModule)

    gcnOracle.CommitTrans
    SaveSet = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Function

Private Sub CmdOK_Click()
    If SaveSet = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    
    If cbo票据设置.ListIndex < 0 Then
        ShowMsgBox "请设置好票据!"
        cbo票据设置.SetFocus
    End If
    Select Case cbo票据设置.ListIndex
    Case 0
        '单据打印
        strBill = "ZL1_BILL_1723"
    Case 1
        '清单打印
        strBill = "ZL1_BILL_1723_1"
    Case 2
        '处方退料通知单
        strBill = "ZL1_BILL_1723_2"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Load()
    Dim strReg As String
    Dim i As Long
    Dim strArr As Variant
    Dim str病区发料 As String
    Dim BlnSelect As Boolean
    Dim n As Integer
    Dim int收费处方 As Integer
    
    mblnHavePriv = IsHavePrivs(mstrPrivs, "参数设置")
    
    With cbo收费处方
        .Clear
        .AddItem "1-显示所有的处方"
        .AddItem "2-仅显示已收费处方"
        .AddItem "3-仅显示未收费处方"
        .ListIndex = 0
    End With
    
    With cbo票据设置
        .Clear
        .AddItem "1-卫材处方单"
        .AddItem "2-打印已发料清单"
        .AddItem "3-退料通知单打印"
        .ListIndex = 0
    End With
    
    With Me.cbo发料后
        .AddItem "1-发料后提示是否打印"
        .AddItem "2-发料后自动打印"
        .AddItem "3-发料后不打印"
        .ListIndex = 0
    End With
    
    With Me.cbo退料后
        .AddItem "1-退料后提示是否打印"
        .AddItem "2-退料后自动打印"
        .AddItem "3-退料后不打印"
        .ListIndex = 0
    End With
    
    chk销帐.Value = IIf(Val(zlDatabase.GetPara("自动销帐", glngSys, mlngModule, , Array(chk销帐), mblnHavePriv)) = 1, 1, 0)
  
    Chk是否自动缺料检查.Value = IIf(Val(zlDatabase.GetPara("缺料检查", glngSys, mlngModule, , Array(Chk是否自动缺料检查), mblnHavePriv)) = 1, 1, 0)
    str病区发料 = zlDatabase.GetPara("病区发料方式", glngSys, mlngModule, "", Array(chk病区, chkDeptType(0), chkDeptType(1), chkDeptType(2), chkDeptType(3), chkDeptType(4), chkDeptType(5), chkDeptType(6)), mblnHavePriv)
        
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0", Array(chk单位(0), chk单位(1), fra(0)), mblnHavePriv))
    chk单位(0).Value = 0
    chk单位(1).Value = 0
    If Val(strReg) = 0 Then
        chk单位(0).Value = 1
    Else
        chk单位(1).Value = 1
    End If
      
    strReg = Trim(zlDatabase.GetPara("发料打印提醒方式", glngSys, mlngModule, "0", Array(cbo发料后), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        cbo发料后.ListIndex = Val(strReg)
    Else
        cbo发料后.ListIndex = 0
    End If
    
    strReg = Trim(zlDatabase.GetPara("退料打印提醒方式", glngSys, mlngModule, "0", Array(cbo退料后), mblnHavePriv))
    If Val(strReg) >= 0 And Val(strReg) <= 2 Then
        cbo退料后.ListIndex = Val(strReg)
    Else
        cbo退料后.ListIndex = 0
    End If
    
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, "", Array(lbl单据类型, chk业务(0), chk业务(1), chk业务(2), Frame3), mblnHavePriv))
    If strReg = "" Then strReg = "24,25,26"
    strArr = Split(strReg & "," & "," & ",", ",")
    For i = 0 To UBound(strArr)
        If i > 2 Then Exit For
        chk业务(i).Value = IIf(Val(strArr(i)) > 0, 1, 0)
    Next
    
    chkSendByNo.Value = IIf(Val(zlDatabase.GetPara("按单据号发料", glngSys, mlngModule, , Array(chkSendByNo), mblnHavePriv)) = 1, 1, 0)
    
    '病区发药
    BlnSelect = False
    If str病区发料 = "" Then
        BlnSelect = False
    ElseIf str病区发料 = mstrAllType Then
        BlnSelect = True
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
    Else
        str病区发料 = str病区发料 & ","
        strArr = Split(str病区发料, ",")
        
        For n = 0 To chkDeptType.Count - 1
            chkDeptType(n).Value = 1
        Next
        
        For i = 0 To UBound(strArr)
            For n = 0 To chkDeptType.Count - 1
                If strArr(i) = chkDeptType(n).Caption Then
                    chkDeptType(n).Value = 0
                    BlnSelect = True
                    Exit For
                End If
            Next
        Next
    End If
    If BlnSelect = True Then
        chk病区.Value = 1
        chk病区.Tag = 1
    Else
        chk病区.Value = 0
        chk病区.Tag = 0
    End If
    For n = 0 To chkDeptType.Count - 1
        chkDeptType(n).Enabled = BlnSelect
    Next
    
    int收费处方 = Val(zlDatabase.GetPara("收费处方显示方式", glngSys, mlngModule, 0, Array(lbl收费处方, cbo收费处方), mblnHavePriv))
    If int收费处方 >= 0 And int收费处方 <= 2 Then
        cbo收费处方.ListIndex = int收费处方
    Else
        cbo收费处方.ListIndex = 0
    End If
    
    '59655
    chkSign.Value = IIf(Val(zlDatabase.GetPara("领料人签名", glngSys, mlngModule, , Array(chkSign), mblnHavePriv)) = 1, 1, 0)
    
    chk汇总发料.Value = IIf(Val(zlDatabase.GetPara("发料时汇总退料销帐记录", glngSys, mlngModule, , Array(chk汇总发料), mblnHavePriv)) = 1, 1, 0)
    
    chk发生时间.Value = IIf(Val(zlDatabase.GetPara("卫材医嘱按发生时间过滤", glngSys, mlngModule, 0, Array(chk发生时间), mblnHavePriv)) = 1, 1, 0)
End Sub
 
Public Function ShowSetPara(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:设置参数入口
    '参数:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴宏
    '修改:2007/12/24
    '-----------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    '本地参数设置
     Me.Show 1, frmMain
    ShowSetPara = mblnOk
End Function
