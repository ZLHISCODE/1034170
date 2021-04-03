VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillDiscard 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票据报损"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBillDiscard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton opt范围 
      Caption         =   "多张报损(&M)"
      Height          =   240
      Index           =   1
      Left            =   3300
      TabIndex        =   19
      Top             =   1410
      Width           =   1635
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "单张报损(&S)"
      Height          =   240
      Index           =   0
      Left            =   1545
      TabIndex        =   18
      Top             =   1410
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   1530
      TabIndex        =   6
      Top             =   870
      Width           =   1815
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      Index           =   2
      Left            =   4110
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1875
      Width           =   1485
   End
   Begin VB.TextBox txtEdit 
      Height          =   360
      Index           =   1
      Left            =   1860
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1875
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   360
      Left            =   4875
      TabIndex        =   10
      Top             =   2430
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   110690307
      CurrentDate     =   37007
   End
   Begin VB.ComboBox cmb报损人 
      Height          =   360
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2415
      Width           =   1830
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   3630
      TabIndex        =   13
      Top             =   5430
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   4980
      TabIndex        =   14
      Top             =   5430
      Width           =   1200
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
      Height          =   30
      Left            =   -270
      TabIndex        =   12
      Top             =   5160
      Width           =   7065
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   270
      TabIndex        =   15
      Top             =   5430
      Width           =   1200
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Index           =   2
      Left            =   3780
      TabIndex        =   17
      Top             =   1875
      Width           =   315
   End
   Begin VB.Label lbl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   360
      Index           =   1
      Left            =   1530
      TabIndex        =   16
      Top             =   1875
      Width           =   315
   End
   Begin VB.Label lbl说明 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   2085
      Left            =   150
      TabIndex        =   11
      Top             =   2940
      Width           =   6300
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "～"
      Height          =   240
      Index           =   5
      Left            =   3450
      TabIndex        =   5
      Top             =   1935
      Width           =   240
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "号码范围(&B)"
      Height          =   240
      Index           =   6
      Left            =   150
      TabIndex        =   2
      Top             =   1935
      Width           =   1320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据种类"
      Height          =   240
      Index           =   4
      Left            =   510
      TabIndex        =   1
      Top             =   930
      Width           =   960
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报损时间(&D)"
      Height          =   240
      Index           =   3
      Left            =   3495
      TabIndex        =   9
      Top             =   2490
      Width           =   1320
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "报损人(&G)"
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   7
      Top             =   2475
      Width           =   1080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "票据报损卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2130
      TabIndex        =   0
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "frmBillDiscard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnChange As Boolean     '为真时表示已改变了
Dim mdatCurrnet As Date
Dim mstrID As String
Dim mstr前缀 As String
Dim mstr最小号码 As String
Dim mstr最大号码 As String
Private mstrPrivs As String
Private mlng票据长度 As Long

Private Sub InitContext()
    dtpDate.Value = mdatCurrnet
    dtpDate.MaxDate = mdatCurrnet
    
    txtEdit(0).Text = frmBillSupervise.lvwMain.SelectedItem.Text
    txtEdit(0).Tag = Mid(frmBillSupervise.lvwMain.SelectedItem.Key, 2)

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    mblnChange = True
    Select Case txtEdit(0).Tag
        Case 1      '1-收费收据
            gstrSQL = " And B.人员性质='门诊收费员'"
        Case 2      '2-预交收据
            gstrSQL = " And B.人员性质 in ('预交收款员','入院登记员')"
        Case 3      '3-结帐收据
            gstrSQL = " And B.人员性质='住院结帐员'"
        Case 4      '4-挂号收据
            gstrSQL = " And B.人员性质='门诊挂号员'"
        Case 5      '5-就诊卡
            gstrSQL = " And B.人员性质 in ('发卡登记人','入院登记员')"
        Case Else
            Exit Sub
    End Select
    gstrSQL = "Select A.姓名 From 人员表 A,人员性质说明 B Where A.ID=B.人员ID " & gstrSQL & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) order by A.姓名"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    cmb报损人.Clear
    Do Until rsTemp.EOF
        cmb报损人.AddItem rsTemp("姓名")
        rsTemp.MoveNext
    Loop
    If cmb报损人.ListCount > 0 Then cmb报损人.ListIndex = 0

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmb报损人_Click()
    mblnChange = True
End Sub

Private Sub cmb报损人_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub

Private Sub dtpDate_Change()
    mblnChange = True
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If ValidateContent() = False Then Exit Sub
    If MsgBox("票据一旦报损后，报损号码就不能再使用了。" & vbCrLf & "是否确认要继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If Save() = False Then Exit Sub
    
    '修改
    Call frmBillSupervise.ShowItem(frmBillSupervise.lvw领用_S.SelectedItem)
    frmBillSupervise.Fill汇总
    mblnChange = False
    Unload Me
End Sub

Private Sub opt范围_Click(Index As Integer)
    mblnChange = True
    If opt范围(0).Value = True Then
        txtEdit(2).Enabled = False
        txtEdit(2).Text = txtEdit(1).Text
    Else
        txtEdit(2).Enabled = True
    End If
    Call ShowSum
End Sub

Private Sub opt范围_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 And opt范围(0).Value = True Then txtEdit(2).Text = txtEdit(1).Text
    Call ShowSum
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    SelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
        Exit Sub
    End If
    If (Index = 1 Or Index = 2) And (KeyAscii >= vbKey0 Or KeyAscii <= vbKey9) And txtEdit(Index).SelLength = 0 Then
        If Len(txtEdit(Index)) >= mlng票据长度 Then KeyAscii = 0
    End If
End Sub

Private Function ValidateContent() As Boolean
'功能:检查输入内容的是否有效
'返回:有效则返回True,否则返回False
    Dim lngCount As Long, i As Integer
    Dim strTemp As String
    
    ValidateContent = False
    '字符串检查
    For lngCount = 1 To 2
        txtEdit(lngCount).Text = Trim(txtEdit(lngCount).Text)
        If StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            txtEdit(lngCount).SetFocus
            SelAll txtEdit(lngCount)
            Exit Function
        End If
        For i = 1 To Len(txtEdit(lngCount).Text)
            strTemp = Mid(txtEdit(lngCount), i, 1)
            If InStr("0123456789", strTemp) = 0 Then
                MsgBox "号码中含有非数字字符。", vbExclamation, gstrSysName
                txtEdit(lngCount).SetFocus
                SelAll txtEdit(lngCount)
                Exit Function
            End If
        Next
        If Len(txtEdit(lngCount).Text) <> Len(txtEdit(lngCount).Tag) - Len(mstr前缀) Then
            MsgBox "号码的长度不对。", vbExclamation, gstrSysName
            txtEdit(lngCount).SetFocus
            SelAll txtEdit(lngCount)
            Exit Function
        End If
    Next
    
    If mstr前缀 & txtEdit(1).Text < txtEdit(1).Tag Then
        MsgBox "作废的开始号码必须大于领用的开始号码。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        SelAll txtEdit(1)
        Exit Function
    End If
    If txtEdit(2).Enabled = True Then
        If mstr前缀 & txtEdit(2).Text > txtEdit(2).Tag Then
            MsgBox "作废的终止号码必须小于领用的终止号码。", vbExclamation, gstrSysName
            txtEdit(2).SetFocus
            SelAll txtEdit(2)
            Exit Function
        End If
    Else
        If mstr前缀 & txtEdit(1).Text > txtEdit(2).Tag Then
            MsgBox "作废的号码必须小于领用的终止号码。", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            SelAll txtEdit(1)
            Exit Function
        End If
    End If
        
    If txtEdit(1).Text > txtEdit(2).Text Then
        MsgBox "作废的开始号码必须小于作废的终止号码。", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        SelAll txtEdit(1)
        Exit Function
    End If
    If Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 > 10000 Then
        MsgBox "一次作废的总张数不能超过一万张。", vbExclamation, gstrSysName
        txtEdit(2).SetFocus
        SelAll txtEdit(2)
        Exit Function
    End If
    If mstr最小号码 <> "" Then
        If mstr前缀 & txtEdit(1).Text <= mstr最小号码 And mstr最小号码 <= mstr前缀 & txtEdit(2).Text Or _
                mstr前缀 & txtEdit(1).Text <= mstr最大号码 And mstr最大号码 <= mstr前缀 & txtEdit(2).Text Then
            MsgBox "作废的号码中包含了已经使用的 。", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            SelAll txtEdit(1)
            Exit Function
        End If
        If mstr前缀 & txtEdit(1).Text > mstr最小号码 And mstr前缀 & txtEdit(2).Text < mstr最大号码 Then
            If MsgBox("作废的号码中可能包含了已经使用的，是否继续？", vbYesNo Or vbQuestion Or vbDefaultButton2, gstrSysName) = vbNo Then
                txtEdit(1).SetFocus
                SelAll txtEdit(1)
                Exit Function
            End If
        End If
    End If
    If cmb报损人.Text = "" Then
        MsgBox "报损人不能为空。", vbExclamation, gstrSysName
        cmb报损人.SetFocus
        Exit Function
    End If
    
    ValidateContent = True
End Function

Private Function Save() As Boolean
'功能:保存编辑的内容
'参数:
'返回值:成功返回True,否则为False
    Dim strTemp As String
    Dim lngID As Long
    
    On Error GoTo errHandle
    Save = False
    
    '修改
    gstrSQL = "zl_票据使用明细_damage(" & mstrID & "," & txtEdit(0).Tag & _
        ",'" & mstr前缀 & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & _
        "',to_date('" & Format(dtpDate.Value, "yyyy-MM-dd") & "','yyyy-mm-dd'),'" & cmb报损人.Text & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    If gblnBillPrint Then
        Call gobjBillPrint.zlDiscardBill(mstrID, Val(txtEdit(0).Tag), mstr前缀, txtEdit(1).Text, txtEdit(2).Text, dtpDate.Value, cmb报损人.Text)
    End If
    
    Save = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ShowSum()
'功能:显示汇总信息
    Dim strTemp As String
    strTemp = " 作废的开始号码：" & lbl(1).Caption & txtEdit(1).Text & vbCrLf
    strTemp = strTemp & "  作废的结束号码：" & lbl(2).Caption & txtEdit(2).Text & vbCrLf
    If txtEdit(1).Text = "" Or txtEdit(2).Text = "" Then
        strTemp = strTemp & "  作废的票据总张数：" & vbCrLf & vbCrLf
    Else
        strTemp = strTemp & "  作废的票据总张数：" & Val(txtEdit(2).Text) - Val(txtEdit(1).Text) + 1 & vbCrLf & vbCrLf
    End If
    strTemp = strTemp & "  领用的开始号码：" & Replace(txtEdit(1).Tag, "&", "&&") & vbCrLf
    strTemp = strTemp & "  领用的结束号码：" & Replace(txtEdit(2).Tag, "&", "&&") & vbCrLf
    If mstr最小号码 <> "" Then
        strTemp = strTemp & "  已经使用的最小号码：" & Replace(mstr最小号码, "&", "&&") & vbCrLf
        strTemp = strTemp & "  已经使用的最大号码：" & Replace(mstr最大号码, "&", "&&") & vbCrLf
    End If
    
    lbl说明.Caption = strTemp
End Sub

Public Function 编辑票据报损(ByVal strPrivs As String, ByVal strID As String) As Boolean
'功能:用来与调用的财务监控窗口进行通讯的程序,用来增加缴款记录
'参数:str缴款人     缴款人的名字
'返回值:编辑成功返回True,否则为False
    Dim rsTemp As New ADODB.Recordset
    Dim dblCount As Double
    
    On Error GoTo errHandle
        
    mstrPrivs = strPrivs
    
    mdatCurrnet = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    mstrID = strID
    
    Call InitContext
    
    gstrSQL = "Select 领用人,前缀文本,开始号码,终止号码,当前号码,使用方式 " & _
            " From 票据领用记录 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    
    mstr前缀 = IIf(IsNull(rsTemp("前缀文本")), "", rsTemp("前缀文本"))
    lbl(1).Caption = Replace(mstr前缀, "&", "&&")
    lbl(2).Caption = lbl(1).Caption
    txtEdit(1).Tag = rsTemp("开始号码")
    txtEdit(2).Text = Mid(rsTemp("终止号码"), Len(mstr前缀) + 1)
    mlng票据长度 = Len(Mid(rsTemp("终止号码"), Len(mstr前缀) + 1))
    txtEdit(2).Tag = rsTemp("终止号码")
    If IsNull(rsTemp("当前号码")) Then
        txtEdit(1).Text = Mid(rsTemp("开始号码"), Len(mstr前缀) + 1)
    Else
        '已经使用，就把最大值加一
        dblCount = Val(Mid(rsTemp("当前号码"), Len(mstr前缀) + 1))
        dblCount = dblCount + 1
        txtEdit(1).Text = Format(dblCount, String(Len(txtEdit(2).Text), "0"))
    End If
    
    On Error Resume Next
    If Val(rsTemp!使用方式) = 2 Then    '共享方式下,只能选择为本操作员:35846
  
        cmb报损人.Text = UserInfo.姓名
    Else
        cmb报损人.Text = rsTemp("领用人")
    End If
    If Err <> 0 Then
        If Val(rsTemp!使用方式) = 2 Then
            cmb报损人.AddItem UserInfo.姓名
            cmb报损人.ListIndex = cmb报损人.NewIndex
        Else
            cmb报损人.AddItem rsTemp("领用人")
            cmb报损人.ListIndex = cmb报损人.NewIndex
        End If
    End If
    If InStr(mstrPrivs, "所有操作员") = 0 Then cmb报损人.Enabled = False
    On Error GoTo errHandle
    
    gstrSQL = "select nvl(min(号码),' ') as 最小号码,nvl(max(号码),' ')  as 最大号码 from 票据使用明细 where 领用ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrID)
    
    mstr最小号码 = Trim(rsTemp("最小号码"))
    mstr最大号码 = Trim(rsTemp("最大号码"))
    Call opt范围_Click(0)
    
    mblnChange = False
    frmBillDiscard.Show vbModal, frmBillSupervise
    编辑票据报损 = True
    Exit Function
errHandle:
    MsgBox "数据读出失败。", vbExclamation, gstrSysName
    编辑票据报损 = False
End Function
