VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeTurnParSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "门诊费用转住院参数设置"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "本地共用预交款票据"
      Height          =   2070
      Left            =   225
      TabIndex        =   11
      Top             =   1035
      Width           =   5040
      Begin MSComctlLib.ListView lvwDeposit 
         Height          =   1755
         Left            =   90
         TabIndex        =   12
         Top             =   210
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   3096
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img16"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "领用人"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "领用日期"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "号码范围"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "剩余"
            Object.Width           =   1499
         EndProperty
      End
   End
   Begin VB.CheckBox chk立即销帐 
      Caption         =   "门诊费用转住院费用后立即退费或销帐(X)"
      Height          =   180
      Left            =   264
      TabIndex        =   10
      Top             =   4020
      Width           =   4116
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   8
      Top             =   4305
      Width           =   5670
   End
   Begin VB.Frame fraTopSplit 
      Height          =   45
      Left            =   -30
      TabIndex        =   7
      Top             =   855
      Width           =   5670
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4095
      TabIndex        =   6
      Top             =   4515
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2910
      TabIndex        =   5
      Top             =   4515
      Width           =   1100
   End
   Begin VB.Frame fraDeposit 
      Caption         =   "预交款票据"
      ForeColor       =   &H00000000&
      Height          =   750
      Left            =   225
      TabIndex        =   0
      Top             =   3180
      Width           =   5040
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "打印设置"
         Height          =   345
         Index           =   0
         Left            =   3840
         TabIndex        =   4
         Top             =   195
         Width           =   990
      End
      Begin VB.OptionButton optPrepayPrint 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   285
         Value           =   -1  'True
         Width           =   900
      End
      Begin VB.OptionButton optPrepayPrint 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   1140
         TabIndex        =   2
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton optPrepayPrint 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   1
         Top             =   285
         Width           =   1380
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeTurnParSet.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   门诊费用转住院费用的相关参数设置，在设置参数时，请注意各参数的含义。"
      Height          =   375
      Left            =   735
      TabIndex        =   9
      Top             =   330
      Width           =   4380
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   105
      Picture         =   "frmChargeTurnParSet.frx":00E2
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmChargeTurnParSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrPrivs As String, mlngModule As Long
Private mblnOk As Boolean
Public Function ShowSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示和设置
    '返回:成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-16 09:50:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    mstrPrivs = strPrivs: mlngModule = lngModule: mblnOk = False
    Me.Show 1, frmMain
    ShowSet = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1131)
End Sub
 
Private Sub cmdOK_Click()
    Dim i As Integer
    '本地共用预交款票据
    zlDatabase.SetPara "共用预交票据批次", 0, glngSys, mlngModule, True
    For i = 1 To lvwDeposit.ListItems.Count
        If lvwDeposit.ListItems(i).Checked Then
            zlDatabase.SetPara "共用预交票据批次", Mid(lvwDeposit.SelectedItem.Key, 2), glngSys, mlngModule, True
        End If
    Next
        
    '预交款票据打印
    For i = 0 To optPrepayPrint.UBound
        If optPrepayPrint(i).Value Then
            zlDatabase.SetPara "门诊转住院预交打印", i, glngSys, mlngModule, IIf(optPrepayPrint(i).Enabled = True, True, False)
        End If
    Next
    zlDatabase.SetPara "费用转出立即退费", chk立即销帐.Value, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
    mblnOk = True
    Unload Me
End Sub
Private Sub cmdPrintSet_Click(Index As Integer)
    Select Case Index
    Case 0  '预交款打印设置
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub Form_Load()
    Dim i As Integer, blnBill As Boolean
    Dim rsTmp As ADODB.Recordset, objItem As ListItem
    On Error GoTo errH
    '读取可用公用预交领用:    '问题:36984
    Set rsTmp = GetShareInvoiceGroupID(2)
    blnBill = False
    rsTmp.Filter = "使用类别=2"
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objItem = lvwDeposit.ListItems.Add(, "_" & rsTmp!ID, rsTmp!领用人, , 1)
            objItem.SubItems(1) = Format(rsTmp!登记时间, "yyyy-MM-dd")
            objItem.SubItems(2) = rsTmp!开始号码 & "," & rsTmp!终止号码
            objItem.SubItems(3) = rsTmp!剩余数量
            If rsTmp!ID = zlDatabase.GetPara("共用预交票据批次", glngSys, mlngModule) Then
                objItem.Selected = True
                objItem.Checked = True
                blnBill = True
            End If
            rsTmp.MoveNext
        Next
    End If
    If Not blnBill Then zlDatabase.SetPara "共用预交票据批次", 0, glngSys, mlngModule, True
    i = Val(zlDatabase.GetPara("门诊转住院预交打印", glngSys, mlngModule, , Array(fraDeposit, optPrepayPrint(0), optPrepayPrint(1), optPrepayPrint(2)), InStr(mstrPrivs, "参数设置") > 0))
    If i <= optPrepayPrint.UBound Then optPrepayPrint(i).Value = True
    chk立即销帐.Value = IIf(Val(zlDatabase.GetPara("费用转出立即退费", glngSys, mlngModule, , Array(chk立即销帐), InStr(mstrPrivs, "参数设置") > 0)) = 1, 1, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwDeposit_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwDeposit.ListItems.Count
        If lvwDeposit.ListItems(i).Key <> Item.Key Then lvwDeposit.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub
