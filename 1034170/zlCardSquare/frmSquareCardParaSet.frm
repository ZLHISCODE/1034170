VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardParaSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "参数设置"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6792
   Icon            =   "frmSquareCardParaSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   6792
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk连续充值 
      Caption         =   "充值后不退出充值界面(&N)"
      Height          =   240
      Left            =   4215
      TabIndex        =   10
      Top             =   2700
      Width           =   2400
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   4020
      Width           =   9600
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   -60
      Width           =   9600
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   350
      Left            =   4290
      TabIndex        =   7
      Top             =   3345
      Width           =   1500
   End
   Begin VB.Frame fra 
      Caption         =   "消费卡设置"
      Height          =   2220
      Left            =   90
      TabIndex        =   4
      Top             =   195
      Width           =   6525
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   1770
         Left            =   60
         TabIndex        =   5
         Top             =   315
         Width           =   6390
         _cx             =   11271
         _cy             =   3122
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "打印设置"
      Height          =   1245
      Left            =   75
      TabIndex        =   2
      Top             =   2535
      Width           =   4005
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "票据打印设置"
         Height          =   360
         Left            =   555
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   705
         Width           =   1875
      End
      Begin VB.CheckBox chk缴款单 
         Caption         =   "缴款立即打印缴款单"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   270
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4215
      TabIndex        =   1
      Top             =   4290
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5400
      TabIndex        =   0
      Top             =   4290
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareCardParaSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs   As String, mblnFirst As Boolean, mblnChange As Boolean
Private Sub InitSqure()
    Dim strColHead As String, varData As Variant, i As Long
    strColHead = "消费卡接口名称|结算方式|卡号前缀文本|卡号长度|卡号密文显示"
     varData = Split(strColHead, "|")
    With vsCardList
            .Clear
            .Cols = UBound(varData) + 1
            For i = 0 To UBound(varData)
                .FixedAlignment(i) = flexAlignCenterCenter
                .TextMatrix(0, i) = varData(i)
                If varData(i) = "卡号密文显示" Then
                    .ColDataType(i) = flexDTBoolean
                End If
                .ColKey(i) = varData(i)
            Next
    End With
End Sub
Private Function LoadCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载消费卡信息
    '返回:
    '编制:刘兴洪
    '日期:2009-12-15 11:29:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, lngRow As Long
    ' 编码,名称,结算方式,nvl(自制卡,0)  as 自制卡,前缀文本,卡号长度
    
    On Error GoTo errHandle
    
    Set rsTemp = zlGet消费卡接口
    rsTemp.Filter = "自制卡=1"
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "消费卡接口不存在,请检查!"
        Exit Function
    End If
    With vsCardList
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("消费卡接口名称")) = Nvl(rsTemp!编号) & "-" & Nvl(rsTemp!名称)
            .Cell(flexcpData, lngRow, .ColIndex("消费卡接口名称")) = Nvl(rsTemp!编号)
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .Cell(flexcpData, lngRow, .ColIndex("结算方式")) = .TextMatrix(lngRow, .ColIndex("结算方式"))
            .TextMatrix(lngRow, .ColIndex("卡号前缀文本")) = Nvl(rsTemp!前缀文本)
            .Cell(flexcpData, lngRow, .ColIndex("卡号前缀文本")) = .TextMatrix(lngRow, .ColIndex("卡号前缀文本"))
            .TextMatrix(lngRow, .ColIndex("卡号长度")) = Nvl(rsTemp!卡号长度, 20)
            .Cell(flexcpData, lngRow, .ColIndex("卡号长度")) = .TextMatrix(lngRow, .ColIndex("卡号长度"))
            .TextMatrix(lngRow, .ColIndex("卡号密文显示")) = Val(Nvl(rsTemp!是否密文))
            .Cell(flexcpData, lngRow, .ColIndex("卡号密文显示")) = Val(Nvl(rsTemp!是否密文))
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If InStr(1, mstrPrivs, ";参数设置;") > 0 Then
            gstrSQL = "Select distinct a.结算方式 From 结算方式应用 A,结算方式 b Where a.应用场合 in ('收费','结帐') AND a.结算方式=b.名称 and b.性质=8 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            .ColComboList(.ColIndex("结算方式")) = .BuildComboList(rsTemp, "结算方式", "结算方式")
            .Editable = flexEDKbdMouse
        End If
        .Cell(flexcpForeColor, 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    End With
    LoadCardInfor = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ShowParaSet(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置入口
    '入参:frmMain-父窗口
    '     lngModule-模块号
    '     strPrivs-权限串

    '编制:刘兴洪
    '日期:2009-11-19 15:29:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mblnFirst = True
    Me.Show 1, frmMain
End Sub
Private Sub LoadParaSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载参数设置
    '编制:刘兴洪
    '日期:2009-12-10 17:03:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, varData As Variant
    Dim blnIsHavePriv As Boolean
    blnIsHavePriv = InStr(1, mstrPrivs, ";参数设置;") > 0
    chk缴款单.Value = IIf(Val(zlDatabase.GetPara("缴款单打印", glngSys, mlngModule, , Array(chk缴款单), blnIsHavePriv)) = 1, 1, 0)
    chk连续充值.Value = IIf(Val(zlDatabase.GetPara("连续充值", glngSys, mlngModule, , Array(chk连续充值), blnIsHavePriv)) = 1, 1, 0)
    Call LoadCardInfor
End Sub

 

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:参数设置检查
    '编制:刘兴洪
    '日期:2009-12-10 17:15:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsHavePriv As Boolean, lngLen As Long, rsTemp As ADODB.Recordset, strTemp As String
    Dim lngRow  As Long
    
    On Error GoTo errHandle
    
    blnIsHavePriv = InStr(1, mstrPrivs, ";参数设置;") > 0

    If blnIsHavePriv Then
        With vsCardList
            For lngRow = 1 To .Rows - 1
                strTemp = Trim(.TextMatrix(lngRow, .ColIndex("卡号前缀文本")))
                lngLen = zlCommFun.ActualLen(Trim(strTemp)) + Val(.TextMatrix(lngRow, .ColIndex("卡号长度")))
                If lngLen > 20 Then
                    ShowMsgbox "消费卡号的最大长度(前缀+卡号长度)不能大于20位,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), "|") > 0 Then
                    ShowMsgbox "消费卡号的前缀文本中不能包含:“|,'～~;”,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                
                If InStr(1, Trim(strTemp), ",") > 0 Then
                    ShowMsgbox "消费卡号的前缀文本中不能包含:“|,'～~;”,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), ";") > 0 Then
                    ShowMsgbox "消费卡号的前缀文本中不能包含:“|,';”,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                
                If InStr(1, Trim(strTemp), "'") > 0 Then
                    ShowMsgbox "消费卡号的前缀文本中不能包含:“|,'～~;”,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                If InStr(1, Trim(strTemp), "～") > 0 Or InStr(1, Trim(strTemp), "~") > 0 Then
                    ShowMsgbox "消费卡号的前缀文本中不能包含:“|,'～~;”,请检查"
                    .Row = lngRow
                    zl_CtlSetFocus vsCardList
                    Exit Function
                End If
                        
                If Val(.Cell(flexcpData, lngRow, .ColIndex("卡号长度"))) <> Val(.TextMatrix(lngRow, .ColIndex("卡号长度"))) Or Len(.Cell(flexcpData, lngRow, .ColIndex("卡号前缀文本"))) <> Len(strTemp) Then
                    '发生了更改,所以需要检查长度是否改小
                    gstrSQL = "Select 1 From 消费卡目录 where ID>0 and rownum =1 and 接口编号=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(.Cell(flexcpData, lngRow, .ColIndex("消费卡接口名称"))))
                    If Not rsTemp.EOF Then
                        If lngLen < zlCommFun.ActualLen(Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号前缀文本")))) + Val(.Cell(flexcpData, lngRow, .ColIndex("卡号长度"))) Then
                            ShowMsgbox "由于发生了发卡信息,所以消费卡号的不能调整,请检查"
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    End If
    
    IsValied = True
 
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
        
End Function


Private Function SaveSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向数据库保存参数设置
    '返回:保存成功返回True,否则返回False
    '编制:刘兴洪
    '日期:2009-12-10 16:59:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnIsHavePriv As Boolean, lng接口序号 As Long, lngRow As Long
    Dim cllPro As Collection, blnTrans As Boolean
    
    blnIsHavePriv = InStr(1, mstrPrivs, ";参数设置;") > 0
    Err = 0: On Error GoTo ErrHand:
   
    With vsCardList
        If blnIsHavePriv Then
            Set cllPro = New Collection
            For lngRow = 1 To .Rows - 1
                 lng接口序号 = Val(.Cell(flexcpData, lngRow, .ColIndex("消费卡接口名称")))
                 If lng接口序号 <> 0 Then
                    If Trim(.TextMatrix(lngRow, .ColIndex("结算方式"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("结算方式"))) Or _
                       Trim(.TextMatrix(lngRow, .ColIndex("卡号前缀文本"))) <> Trim(.Cell(flexcpData, lngRow, .ColIndex("卡号前缀文本"))) Or _
                       Val(.TextMatrix(lngRow, .ColIndex("卡号长度"))) <> Val(.Cell(flexcpData, lngRow, .ColIndex("卡号长度"))) Or Abs(Val(.TextMatrix(lngRow, .ColIndex("卡号密文显示")))) <> Val(.Cell(flexcpData, lngRow, .ColIndex("卡号密文显示"))) Then
                           '只有发生了改变的才能更新
                           ' Zl_卡消费接口目录_Update
                           gstrSQL = "Zl_卡消费接口目录_Update("
                           '  编号_In     In 卡消费接口目录.编号%Type,
                           gstrSQL = gstrSQL & "" & lng接口序号 & ","
                           '  结算方式_In In 卡消费接口目录.结算方式%Type,
                           gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("结算方式")) & "',"
                           '  卡号前缀_In In 卡消费接口目录.卡号前缀%Type,
                           gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("卡号前缀文本")) & "',"
                           '  卡号长度_In In 卡消费接口目录.卡号长度%Type
                           gstrSQL = gstrSQL & " " & Val(.TextMatrix(lngRow, .ColIndex("卡号长度"))) & ","
                           '    是否密文_In In 卡消费接口目录.是否密文%Type := 0
                           gstrSQL = gstrSQL & " " & IIf(Abs(Val(.TextMatrix(lngRow, .ColIndex("卡号密文显示")))) = 0, 0, 1) & ")"
                           zlAddArray cllPro, gstrSQL
                    End If
                 End If
            Next
        End If
    End With
    gcnOracle.BeginTrans
    blnTrans = True
    If Not cllPro Is Nothing Then
        If cllPro.Count > 0 Then zlExecuteProcedureArrAy cllPro, Me.Caption, True, blnTrans
    End If
    Call zlDatabase.SetPara("缴款单打印", IIf(chk缴款单.Value = 1, 1, 0), glngSys, mlngModule, blnIsHavePriv)
    Call zlDatabase.SetPara("连续充值", IIf(chk连续充值.Value = 1, 1, 0), glngSys, mlngModule, blnIsHavePriv)
    gcnOracle.CommitTrans: blnTrans = False
    Set grsStatic.rs消费卡接口 = Nothing
    Call zlGet消费卡接口
    SaveSet = True
    Exit Function
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    SaveErrLog
End Function

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, glngSys, mlngModule)
End Sub

Private Sub CmdOK_Click()
    If IsValied = False Then Exit Sub
    If SaveSet = False Then Exit Sub
    Unload Me
End Sub

Private Sub cmdPrintSet_Click()
    Dim strBill As String
    strBill = "ZL1_BILL_1503"
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadCardInfor = False Then Unload Me: Exit Sub
    Call LoadParaSet
End Sub

 

Private Sub Form_Load()
    Call InitSqure
End Sub

Private Sub vsCardList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCardList
        Select Case Col
        Case .ColIndex("结算方式")
        Case .ColIndex("卡号前缀文本")
        Case .ColIndex("卡号长度")
        Case .ColIndex("卡号密文显示")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsCardList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsCardList
        Select Case Col
        Case .ColIndex("结算方式"), .ColIndex("卡号前缀文本"), .ColIndex("卡号长度"), .ColIndex("卡号密文显示")
        Case Else
            Exit Sub
        End Select
    End With
End Sub
Private Sub vsCardList_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsCardList_EnterCell()
    
    With vsCardList
        .EditMaxLength = 0
        Select Case .Col
        Case .ColIndex("卡号前缀文本")
            .EditMaxLength = 4
        Case .ColIndex("卡号长度")
            .EditMaxLength = 3
        End Select
    End With
End Sub

Private Sub vsCardList_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsCardList
        Select Case Col
        Case .ColIndex("结算方式"), .ColIndex("卡号前缀文本"), .ColIndex("卡号长度")
        Case Else
        End Select
        Call zlVsMoveGridCell(vsCardList, 0, .Cols - 1, False)
    End With
End Sub

Private Sub vsCardList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsCardList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    
    With vsCardList
        Select Case Col
        Case .ColIndex("卡号前缀文本")
            If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Case .ColIndex("卡号长度")
            '主要可能存在退款情况
            Call VsFlxGridCheckKeyPress(vsCardList, Row, Col, KeyAscii, m金额式)
        Case Else
        End Select
    End With
End Sub
 
