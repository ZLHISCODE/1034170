VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmChildStationDrug 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   2145
      Index           =   0
      Left            =   285
      TabIndex        =   0
      Top             =   675
      Width           =   3990
      _cx             =   7038
      _cy             =   3784
      Appearance      =   3
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483626
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmChildStationDrug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'（１）窗体级变量定义

Private mblnDataChanged As Boolean
Private mblnReading As Boolean
Private mfrmMain As Object
Private mblnAllowModify As Boolean
Private mlngKey As Long
Private mbytMode As Byte
Private mstrPrivs As String
Private mstr病人来源 As String
Private mlng病人科室id As Long
Private mlng医嘱id As Long
Private mobjStateInfo As CommandBarControl
Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1
Private mint病人来源 As Integer
Public Event AfterDataChanged()
Public Event AfterMakeCharge()

'######################################################################################################################
'（２）自定义过程或函数

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    Set mfrmMain = frmMain
        
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, _
                            Optional ByVal blnAllowModify As Boolean = True, _
                            Optional ByVal bytMode As Byte = 1, _
                            Optional ByVal str病人来源 As String, _
                            Optional ByVal lng医嘱id As Long, _
                            Optional ByVal strPrivs As String) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mbytMode = bytMode
    mlng医嘱id = lng医嘱id
    mstr病人来源 = str病人来源
    mint病人来源 = IIf(str病人来源 = "住院", 2, 1)
    
    mstrPrivs = strPrivs
    mblnAllowModify = blnAllowModify
    mlngKey = lngKey
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    
    If mlngKey > 0 Then
        If ExecuteCommand("读取数据") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If lngLoop <> .Rows - 1 Then
                If Val(.RowData(lngLoop)) = 0 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，必须输入有效的药品项目！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("用药类型"))
                    Exit Function
                End If
            End If
            
            If Val(.RowData(lngLoop)) > 0 Then
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("准备数量"))) = False And mbytMode = 1 And .TextMatrix(lngLoop, .ColIndex("准备数量")) <> "" Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，药品的准备数量必须为数值型！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("准备数量"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("准备数量"))) < 0 And mbytMode = 1 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，必须输入药品的准备数量[0-99999999]！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("准备数量"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("准备数量"))) > 999999999 And mbytMode = 1 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据太大，输入[0-999999999]范围内的数值！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("准备数量"))
                    Exit Function
                End If
                
                If IsNumeric(.TextMatrix(lngLoop, .ColIndex("实用数量"))) = False And mbytMode = 2 And .TextMatrix(lngLoop, .ColIndex("实用数量")) <> "" Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，药品的实用数量必须为数值型！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("实用数量"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("实用数量"))) < 0 And mbytMode = 2 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据输入不完整，必须输入药品的实用数量[0-99999999]！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("实用数量"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("实用数量"))) > 99999999 And mbytMode = 2 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行数据太大，输入[0-999999999]范围内的数值！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("实用数量"))
                    Exit Function
                End If
                
                If Val(.TextMatrix(lngLoop, .ColIndex("执行科室id"))) <= 0 Then
                    ShowSimpleMsg "第 " & lngLoop & " 行没有指定执行科室！"
                    Call LocationGrid(vsf(0), lngLoop, .ColIndex("执行科室id"))
                    Exit Function
                End If
            End If
        Next
    End With
    
    ValidData = True
    
End Function

Public Function ClearData() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Call ExecuteCommand("清空数据")
    
    ClearData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_病人手术用药_DELETE(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    
    
    With vsf(0)
        For lngLoop = 1 To .Rows - 1
            If .RowData(lngLoop) > 0 Then
                strSQL = "ZL_病人手术用药_INSERT(" & mlngKey & ",'" & zlCommFun.GetNeedName(.TextMatrix(lngLoop, .ColIndex("用药类型"))) & "'," & lngLoop & "," & Val(.RowData(lngLoop)) & ",'" & .TextMatrix(lngLoop, .ColIndex("药品名称")) & "'," & Val(.TextMatrix(lngLoop, .ColIndex("准备数量"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("实用数量"))) & "," & Val(.TextMatrix(lngLoop, .ColIndex("执行科室id"))) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With

    SaveData = True
    
    Exit Function
    
    '
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    
    Dim cbrCustom As CommandBarControlCustom
    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)
    
    cbsMain.Options.LargeIcons = False
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "增加", , , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除", , , xtpButtonIconAndCaption)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Default, "缺省", , conMenu_Edit_Modify, xtpButtonIconAndCaption)
    Set objPopup = NewToolBar(objBar, xtpControlPopup, conMenu_Edit_MakeCharge, "生成", True, , xtpButtonIconAndCaption)
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Preferences, "方案", True, , xtpButtonIconAndCaption)

    Set mobjStateInfo = NewToolBar(objBar, xtpControlLabel, 0, "")
    mobjStateInfo.Flags = xtpFlagRightAlign
        
End Function

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim objArray As Variant
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    Call SQLRecord(rsSQL)
    
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "初始控件"

        Set mclsVsf = New clsVsf
        With mclsVsf
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("用药类型", 1080, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("药品名称", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            
            Call .AppendColumn("规格", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("准备数量", 900, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("实用数量", 900, flexAlignRightCenter, flexDTString, "0.00", , True)
            Call .AppendColumn("单位", 600, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("执行科室", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("执行科室id", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("类别", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("可用库存", 0, flexAlignLeftCenter, flexDTString, "", , True, , , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
                        
            .AppendRows = True
        End With
        
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
    
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
    
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        cbsMain.RecalcLayout
        
    '--------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify And (IsPrivs(mstrPrivs, "用药准备") Or IsPrivs(mstrPrivs, "用药登记"))
        If mlngKey = 0 Then blnAllowModify = False
        
        With mclsVsf

            If blnAllowModify Then
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("用药类型"), True, vbVsfEditCombox, " |")
                Call .InitializeEditColumn(.ColIndex("药品名称"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("执行科室"), True, vbVsfEditCombox)
                
                If mbytMode = 1 Then
                    Call .InitializeEditColumn(.ColIndex("准备数量"), True, vbVsfEditText)
                    Call .InitializeEditColumn(.ColIndex("实用数量"), False, vbVsfEditText)
                Else
                    Call .InitializeEditColumn(.ColIndex("准备数量"), False, vbVsfEditText)
                    Call .InitializeEditColumn(.ColIndex("实用数量"), True, vbVsfEditText)
                End If
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
                
                gstrSQL = "Select 编码||'-'||名称 As 名称 From 手术用药类型 Order By 编码"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                If rs.BOF = False Then
                    .DropColData(.ColIndex("用药类型")) = vsf(0).BuildComboList(rs, "名称", "名称", RGB(255, 255, 255))
                End If
                
            Else
                Call .InitializeEdit(False, False, False)
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            End If
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
            
        mclsVsf.ClearGrid
        mobjStateInfo.Caption = " "
        
        gstrSQL = "SELECT a.序号," & _
                        "f.计算单位 As 单位," & _
                        "a.类型 As 用药类型," & _
                        "a.药品ID As ID," & _
                        "a.药品名称," & _
                        "a.准备总量 AS 准备数量," & _
                        "a.使用总量 AS 实用数量," & _
                        "D.名称 AS 执行科室," & _
                        "a.执行科室id,F.类别," & _
                        "f.规格 " & _
                "FROM 病人手术用药 A,药品规格 B,病人手术记录 C,部门表 D,收费项目目录 F " & _
                "WHERE A.药品id=B.药品id " & _
                    "AND B.药品id=F.ID " & _
                    "AND A.执行科室id=D.ID And C.ID=a.记录id " & _
                    "AND A.记录id=[1] " & _
                "ORDER BY A.序号"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsf.LoadGrid(rs)
                
        mobjStateInfo.Caption = ""
        gstrSQL = "Select No,记录性质 From 病人手术单据 Where 记录id=[1] And 单据类型=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 1)
        If rs.BOF = False Then
            If IsNull(rs("No").Value) = False Then
                Select Case rs("记录性质").Value
                Case 1
                    mobjStateInfo.Caption = "已生成收费单，单据号：" & rs("No").Value
                Case 2
                    mobjStateInfo.Caption = "已生成记帐单，单据号：" & rs("No").Value
                End Select
            End If
        End If
        cbsMain.RecalcLayout
    '------------------------------------------------------------------------------------------------------------------
    Case "收费执行科室"
        
        With vsf(0)
        
            gstrSQL = GetPublicSQL(SQL.收费执行科室, CStr(varParam(0)))
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetDefaultDept(CStr(varParam(0)), mint病人来源), Val(.RowData(.Row)), mlng病人科室id, UserInfo.部门ID)
            
            If rs.BOF = False Then
                .TextMatrix(.Row, .ColIndex("执行科室")) = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(.Row, .ColIndex("执行科室id")) = zlCommFun.NVL(rs("ID").Value)
                
                .ColComboList(.ColIndex("执行科室")) = .BuildComboList(rs, "名称", "ID")
                
            Else
                .TextMatrix(.Row, .ColIndex("执行科室")) = UserInfo.部门名称
                .TextMatrix(.Row, .ColIndex("执行科室id")) = UserInfo.部门ID
                .ColComboList(.ColIndex("执行科室")) = " |"
            End If
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case "药品参考方案"
        
        With vsf(0)
            gstrSQL = GetPublicSQL(SQL.手术用药选择)

            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, mlngKey)
            
            If ShowPubSelect(Me, Nothing, 3, "类型,990,0,;编码,900,0,;名称,2400,0,;规格,900,0,;总量,1200,2,;单位,810,0,", mfrmMain.Name & "\手术用药选择", "请从下面左边列表中选择手术用药参考", rsData, rs, 8790, 4500, False, , , True) = 1 Then
                
                gstrSQL = GetPublicSQL(SQL.方案用药参考)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(rs("ID").Value))
                If rs.BOF = False Then
                    mclsVsf.ClearGrid
                    
                    Do While Not rs.EOF
                        If Val(.RowData(.Rows - 1)) > 0 Then .Rows = .Rows + 1
                        .Row = .Rows - 1
                        Call mclsVsf.LoadGridRow(.Row, rs)

                        Call ExecuteCommand("收费执行科室", zlCommFun.NVL(rs("类别").Value))
                        Call CheckStorage(.Row)
                        
                        rs.MoveNext
                    Loop
                    DataChanged = True
                End If
            End If
        
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "生成药品收费单"
    
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要将手术用药生成收费单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlng医嘱id, 1, "用药", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "用药收费单已经生成，单据号：" & strTmp
            mobjStateInfo.Caption = "已生成收费单，单据号：" & strTmp
            cbsMain.RecalcLayout
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "生成药品记帐单"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要将手术用药生成记帐单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlng医嘱id, 2, "用药", False, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "用药记帐单已经生成，单据号：" & strTmp
            mobjStateInfo.Caption = "已生成记帐单，单据号：" & strTmp
            cbsMain.RecalcLayout
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
    '------------------------------------------------------------------------------------------------------------------
    Case "生成药品零费单"
        If DataChanged Then
            If SaveData(rsSQL) Then
                If SQLRecordExecute(rsSQL, mfrmMain.Caption) = False Then Exit Function
                DataChanged = False
            End If
        End If
        
        If MsgBox("如果已生成，则会自动删除或作废，确定要将手术用药生成零费单吗（新单）？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Function
        
        strTmp = MakeChargeBill(mlng医嘱id, 2, "用药", True, mstrPrivs)
        If strTmp <> "" Then
            ShowSimpleMsg "用药零费单已经生成，单据号：" & strTmp
            mobjStateInfo.Caption = "已生成记帐单，单据号：" & strTmp
            cbsMain.RecalcLayout
            RaiseEvent AfterMakeCharge
        End If
        
        Exit Function
        
    End Select
    
    ExecuteCommand = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function CheckStorage(ByVal intRow As Integer) As Boolean
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim byt检查方式 As Byte
    
    With vsf(0)
            
        gstrSQL = GetPublicSQL(SQL.出库检查方式, .TextMatrix(intRow, .ColIndex("类别")))
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, .TextMatrix(intRow, .ColIndex("类别")))
        If rs.BOF = False Then byt检查方式 = zlCommFun.NVL(rs("检查方式").Value, 0)
        If byt检查方式 <> 0 Then
        
            byt检查方式 = zlCommFun.NVL(rs("检查方式").Value, 0)
            
            gstrSQL = GetPublicSQL(SQL.库存数量检查)
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(intRow)), Val(.TextMatrix(intRow, .ColIndex("执行科室id"))))
            If rs.BOF = False Then .TextMatrix(intRow, .ColIndex("可用库存")) = zlCommFun.NVL(rs("库存").Value, 0)
            
            Call PromptStorageWarn(Val(.TextMatrix(intRow, .ColIndex("准备数量"))), Val(.TextMatrix(intRow, .ColIndex("可用库存"))), .TextMatrix(intRow, .ColIndex("药品名称")), .TextMatrix(intRow, .ColIndex("执行科室")), .TextMatrix(intRow, .ColIndex("单位")), byt检查方式)
            
        End If
    End With
    
    CheckStorage = True
End Function

Public Property Get Body(ByVal lngIndex As Long) As Object
    Set Body = vsf
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intRow As Integer
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Preferences                       '方案
    
        Call ExecuteCommand("药品参考方案")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem                           '新增
        
        Call mclsVsf.AppendRow
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete                            '删除
        
        Call mclsVsf.DeleteRow(vsf(0).Row)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Default                           '缺省

        With vsf(0)
            For intRow = 1 To .Rows - 1
                .Cell(flexcpText, intRow, .ColIndex("实用数量")) = .Cell(flexcpText, intRow, .ColIndex("准备数量"))
            Next
            DataChanged = True
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 1

        Call ExecuteCommand("生成药品收费单")

    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 2

        Call ExecuteCommand("生成药品记帐单")
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2# + 3
                
        Call ExecuteCommand("生成药品零费单")

    End Select
End Sub

Private Sub cbsMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl

    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_MakeCharge

        With CommandBar.Controls

            .DeleteAll
            
            
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 1, "收费单据(&1)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 2, "记帐单据(&2)")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MakeCharge * 2 + 3, "零耗费用(&3)")
            With cbsMain.KeyBindings
                .Add FCONTROL, vbKeyN, conMenu_Edit_MakeCharge * 2 + 1
                .Add FCONTROL, vbKeyB, conMenu_Edit_MakeCharge * 2 + 2
            End With
            
        End With
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    
    '窗体其它控件Resize处理
    vsf(0).Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    mclsVsf.AppendRows = True
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnAllowModify As Boolean
    
    On Error GoTo errHand
    
    blnAllowModify = mblnAllowModify
    If mlngKey = 0 Then blnAllowModify = False
        
    Select Case Control.ID
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem, conMenu_Edit_Preferences
        
        Control.Visible = IsPrivs(mstrPrivs, "用药准备") Or IsPrivs(mstrPrivs, "用药登记")
        Control.Enabled = blnAllowModify And Control.Visible
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        Control.Visible = IsPrivs(mstrPrivs, "用药准备") Or IsPrivs(mstrPrivs, "用药登记")
        Control.Enabled = blnAllowModify And Control.Visible And Val(vsf(0).RowData(vsf(0).Row)) > 0
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Default
        
        Control.Visible = IsPrivs(mstrPrivs, "用药登记")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge
        
        Control.Visible = IsPrivs(mstrPrivs, "生成附费")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 1
        
        Control.Visible = (mstr病人来源 = "门诊" And IsPrivs(mstrPrivs, "生成附费"))
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_MakeCharge * 2 + 2, conMenu_Edit_MakeCharge * 2 + 3
        
        Control.Visible = IsPrivs(mstrPrivs, "生成附费")
        Control.Enabled = blnAllowModify And mbytMode = 2 And Control.Visible
        
    End Select

errHand:

End Sub

'######################################################################################################################
'（３）窗体及其控件的事件处理

Private Sub Form_Resize()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjStateInfo = Nothing
    Set mclsVsf = Nothing
End Sub

Private Sub mclsVsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    DataChanged = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '编辑处理
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
    
    With vsf(Index)
        Select Case Col

        Case .ColIndex("用药类型")
                    
            If .ComboIndex > -1 Then
                .TextMatrix(Row, Col) = zlCommFun.GetNeedName(.ComboItem(.ComboIndex))
            End If
                
    
        Case .ColIndex("准备数量"), .ColIndex("实用数量")
            
            Call CheckStorage(Row)
            
        Case .ColIndex("执行科室")
            .TextMatrix(Row, .ColIndex("执行科室id")) = .ComboData
            .TextMatrix(Row, .ColIndex("执行科室")) = .Cell(flexcpTextDisplay, Row, .ColIndex("执行科室"))
            
            Call CheckStorage(Row)
                        
        End Select
    End With
    
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strCode As String
    Dim bln麻醉剂 As Boolean
    
    With vsf(0)
        If Col = .ColIndex("药品名称") Then
            
            strCode = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("用药类型")))
            
            gstrSQL = "Select 是否麻醉剂 From 手术用药类型 Where 名称=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strCode)
            If rs.BOF = False Then bln麻醉剂 = (zlCommFun.NVL(rs("是否麻醉剂").Value, 0) = 1)
                
            If bln麻醉剂 Then
                gstrSQL = GetPublicSQL(SQL.麻醉药品选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                bytRet = ShowPubSelect(Me, vsf(0), 3, "编码,1200,0,;名称,2700,0,;规格,900,0,;单位,900,0,", Me.Name & "\麻醉药品选择", "请从下表中选择一个麻醉项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
            Else
                gstrSQL = GetPublicSQL(SQL.药品项目选择)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                bytRet = ShowPubSelect(Me, vsf(0), 3, "编码,1200,0,;名称,2700,0,;规格,900,0,;单位,900,0,", Me.Name & "\药品项目选择", "请从下表中选择一个药品项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
            End If
            
            If bytRet = 1 Then
                If mclsVsf.CheckHave(zlCommFun.NVL(rs("药品ID").Value)) Then
                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                    Exit Sub
                End If
    
                .EditText = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("药品名称")) = zlCommFun.NVL(rs("名称").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("规格")) = zlCommFun.NVL(rs("规格").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("单位")) = zlCommFun.NVL(rs("单位").Value)
                .TextMatrix(Row, mclsVsf.ColIndex("类别")) = zlCommFun.NVL(rs("类别").Value)
                .RowData(Row) = zlCommFun.NVL(rs("药品ID").Value, 0)
    
                Call ExecuteCommand("收费执行科室", zlCommFun.NVL(rs("类别").Value))
                Call CheckStorage(Row)
                
                DataChanged = True
                Call mclsVsf.LocationNextCell
                
            End If
            
            Call mclsVsf.SetFocus(, , True)
        End If
    End With
End Sub

Private Sub vsf_ChangeEdit(Index As Integer)
    With vsf(Index)
        Select Case .Col
        Case .ColIndex("准备数量"), .ColIndex("实用数量")
            .TextMatrix(.Row, .Col) = .EditText
        Case .ColIndex("用药类型")
            If .EditText <> .TextMatrix(.Row, .Col) Then
                .Cell(flexcpText, .Row, 0, .Row, .Cols - 1) = ""
                .RowData(.Row) = 0
            End If
        End Select

    End With
End Sub

Private Sub vsf_ComboDropDown(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim intLoop As Integer
    
    With vsf(Index)
        Select Case Col
        Case .ColIndex("用药类型")
            If .TextMatrix(Row, Col) <> "" Then
                For intLoop = 0 To .ComboCount - 1
                    If zlCommFun.GetNeedName(.ComboItem(intLoop)) = .TextMatrix(Row, Col) Then
                        .ComboIndex = intLoop
                        Exit For
                    End If
                Next
            End If
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim strCode As String
    Dim bln麻醉剂 As Boolean
    
    With vsf(0)
        If KeyCode = vbKeyReturn Then
            
            If Col = .ColIndex("药品名称") Then
                
                If InStr(vsf(0).EditText, "'") > 0 Then
                    KeyCode = 0
                    vsf(0).EditText = ""
                    Exit Sub
                End If

                strCode = zlCommFun.GetNeedName(.TextMatrix(Row, .ColIndex("用药类型")))
                
                gstrSQL = "Select 是否麻醉剂 From 手术用药类型 Where 名称=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strCode)
                If rs.BOF = False Then bln麻醉剂 = (zlCommFun.NVL(rs("是否麻醉剂").Value, 0) = 1)
            
                strText = UCase(vsf(0).EditText)
                bytMode = GetApplyMode(strText)
                strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, "", "%") & strText & "%"
                
                If bln麻醉剂 Then
                    gstrSQL = GetPublicSQL(SQL.麻醉药品过滤, bytMode)
                Else
                    gstrSQL = GetPublicSQL(SQL.药品项目过滤, bytMode)
                End If

                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)

                If ShowPubSelect(Me, vsf(0), 2, "编码,1200,0,;名称,2700,0,;规格,900,0,;单位,900,0,", Me.Name & "\药品项目过滤", "请从下表中选择一个药品项目", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsf.CheckHave(zlCommFun.NVL(rs("药品ID").Value)) Then
                        ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                        Exit Sub
                    End If

                    .EditText = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("药品名称")) = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("单位")) = zlCommFun.NVL(rs("单位").Value)
                    .TextMatrix(Row, .ColIndex("规格")) = zlCommFun.NVL(rs("规格").Value)
                    .TextMatrix(Row, .ColIndex("类别")) = zlCommFun.NVL(rs("类别").Value)
                    
                    .RowData(Row) = zlCommFun.NVL(rs("药品ID").Value, 0)

                    Call ExecuteCommand("收费执行科室", zlCommFun.NVL(rs("类别").Value))
                    Call CheckStorage(Row)
                
                    DataChanged = True
                    Call mclsVsf.LocationNextCell
                Else
                    KeyCode = 0

                    .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                    .EditText = .Cell(flexcpData, Row, Col)
                    .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                End If
                
                Call mclsVsf.SetFocus(, , True)
            Else
                Call mclsVsf.LocationNextCell
            End If
        Else
            DataChanged = True
        End If
    End With

End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPress(KeyAscii)
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsf.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
    End Select
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        Set cbrPopupBar = CopyMenu(mfrmMain.cbsMain, 3)
        cbrPopupBar.ShowPopup
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsf.ValidateEdit(Col, Cancel)
End Sub


