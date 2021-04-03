VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmProcMain 
   Caption         =   "自定义过程管理"
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15990
   Icon            =   "frmProcMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15990
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3555
      Index           =   0
      Left            =   90
      ScaleHeight     =   3555
      ScaleWidth      =   12405
      TabIndex        =   0
      Top             =   1470
      Width           =   12405
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFEBD7&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   15
         ScaleHeight     =   375
         ScaleWidth      =   10980
         TabIndex        =   1
         Top             =   15
         Width           =   10980
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   9150
            TabIndex        =   5
            ToolTipText     =   "请直接按回车键进行过滤"
            Top             =   45
            Width           =   1695
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "用户过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   2835
            TabIndex        =   4
            Top             =   90
            Width           =   1305
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "空白过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1455
            TabIndex        =   3
            Top             =   90
            Width           =   1305
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEBD7&
            Caption         =   "变动过程"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   90
            Value           =   -1  'True
            Width           =   1305
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "定位："
            Height          =   180
            Left            =   8625
            TabIndex        =   6
            Top             =   90
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1755
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   1935
         _cx             =   3413
         _cy             =   3096
         Appearance      =   1
         BorderStyle     =   0
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
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   9975
      Width           =   15990
      _ExtentX        =   28205
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmProcMain.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   25329
            MinWidth        =   8819
            Text            =   "当前共有待调整0个；调整中0个"
            TextSave        =   "当前共有待调整0个；调整中0个"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmProcMain.frx":70E6
      Left            =   1080
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmProcMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmProgramEdit As frmProcEdit
Private mfrmDifferenceCheck As frmProcDiffrentCheck
Private mfrmBuildScript As frmProcBuildScript
Private mfrmProcedureRelating As frmProcRelating
Private mfrmCollectUpdate As frmProcCollectUpdate
Private mintProcType As Integer
Private mclsVsf As clsVsf
Private mclsVsfMark As clsVsf
Private mblnReading As Boolean
Private mobjMain As Object

Private mblnStartUp As Boolean

Private Sub InitCommandBar()
    '******************************************************************************************************************
    '功能：初始菜单工具栏
    '参数：无
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objExtendedBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    Set cbsMain.Icons = frmPubIcons.imgPublic.Icons
    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '文件
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.Id = conMenu_FilePopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "退出(&X)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '编辑
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    objMenu.Id = conMenu_EditPopup
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Word, "搜集登记(&S)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "新建登记(&N)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Disuse, "升级完成(&C)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Audit, "差异检查(&J)", True)
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "修改过程(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "删除过程(&D)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Untread, "恢复过程(&R)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Manage_Change_PaitNote, "生成脚本(&G)", True)
    
    '------------------------------------------------------------------------------------------------------------------
    '查看
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.Id = conMenu_ViewPopup
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
    
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "刷新(&R)", True)
    

    '------------------------------------------------------------------------------------------------------------------
    '帮助
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.Id = conMenu_HelpPopup
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
    Set objPopup = gclsBase.NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrWebSustainer)
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, gstrWebSustainer & "主页(&H)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, gstrWebSustainer & "论坛(&F)")
    Set objControl = gclsBase.NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)")
    Set objControl = gclsBase.NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "关于(&A)…", True)
    
    '标准工具栏
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
'    objBar.SetIconSize 16, 16
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Word, "搜集")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "新建")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Audit, "检查", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "修改", True)
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除")
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Edit_Untread, "恢复")
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_Manage_Change_PaitNote, "生成", True)
    
    Set objControl = gclsBase.NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出", True)
    
    '命令的快键绑定:公共部份主界面已处理
    '-----------------------------------------------------
    With cbsMain.KeyBindings
                
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新增
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改
        .Add 0, vbKeyDelete, conMenu_Edit_Delete '删除
                
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
End Sub

Public Function ShowMe(ByVal objParent As Object)
    Me.Show 1, objParent
End Function

Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 600, 100, DockLeftOf, objPane)
    objPane.Title = "过程"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    Call gclsBase.DockPannelInit(dkpMain)

End Sub

Private Function ExecuteCommand(ParamArray varCmd() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim objNode As Node
    Dim intRow As Integer
    Dim rsColum As ADODB.Recordset
    Dim lngKey As Long
    Dim intType1 As Integer
    Dim intType2 As Integer
    Dim intType3 As Integer
    
    On Error GoTo errHand

    Call gclsBase.SQLRecord(rsSQL)

    For intLoop = 0 To UBound(varCmd)
        Select Case varCmd(intLoop)
        '--------------------------------------------------------------------------------------------------------------
        Case "初始控件"
            
            Call InitCommandBar
            Call InitDockPannel
        '--------------------------------------------------------------------------------------------------------------
        Case "初始数据"
            Set mclsVsf = New clsVsf
            With mclsVsf
                Call .Initialize(Me.Controls, vsf(0), True, False)
                Call .ClearColumn
                Call .AppendColumn("", 270, flexAlignLeftCenter, flexDTString, , "[序号]", False, False, False)
'                Call .AppendColumn("序号", 800, flexAlignLeftCenter, flexDTString, , "序号", False)
                Call .AppendColumn("过程", 3000, flexAlignLeftCenter, flexDTString, , "过程", True)
                Call .AppendColumn("状态", 1000, flexAlignLeftCenter, flexDTString, , "状态", True)
                Call .AppendColumn("说明", 3000, flexAlignLeftCenter, flexDTString, , "说明", True)
                
                .IndicatorMode = 2
                .IndicatorCol = .ColIndex("序号")
                .ConstCol = .ColIndex("序号")
            
                .AppendRows = True
            End With
            mintProcType = IIf((opt(0).value = True), 1, IIf((opt(1).value = True), 2, 3))
        '--------------------------------------------------------------------------------------------------------------
        Case "刷新数据"
            
            '清空原有数据
            Call mclsVsf.ClearGrid
            strSQL = "Select ID,Decode(类型,1,'标准过程',2,'空白过程',3,'用户过程') As 类型,名称 As 过程,Decode(状态,1,'待调整',2,'调整中',3,'已调整') As 状态,说明,修改人员,修改时间,上次修改人员,上次修改时间 From zlprocedure Where 类型=[1]"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", mintProcType)
            If rs.BOF = False Then
                Call mclsVsf.LoadGrid(rs)
            End If
            For intRow = 1 To vsf(0).Rows - 1
                If vsf(0).TextMatrix(intRow, vsf(0).ColIndex("状态")) = "待调整" Then
                    vsf(0).Cell(flexcpForeColor, intRow, vsf(0).ColIndex("状态")) = vbRed
                ElseIf vsf(0).TextMatrix(intRow, vsf(0).ColIndex("状态")) = "调整中" Then
                    vsf(0).Cell(flexcpForeColor, intRow, vsf(0).ColIndex("状态")) = vbBlue
                End If
            Next
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        '--------------------------------------------------------------------------------------------------------------
        Case "刷新状态"
            intType1 = 0
            intType2 = 0
            For intRow = 1 To vsf(0).Rows - 1
                If vsf(0).TextMatrix(intRow, vsf(0).ColIndex("状态")) = "待调整" Then
                    intType1 = intType1 + 1
                ElseIf vsf(0).TextMatrix(intRow, vsf(0).ColIndex("状态")) = "调整中" Then
                    intType2 = intType2 + 1
                End If
            Next
            stbThis.Panels(2).Text = "当前共有待调整 " & CStr(intType1) & " 个,调整中 " & CStr(intType2) & " 个。"
        '--------------------------------------------------------------------------------------------------------------
        Case "恢复过程"
            With vsf(0)
                lngKey = .RowData(.Row)
                If lngKey = 0 Then Exit Function
                '得到本次变动过程对应的标准过程
                strSQL = "Select A.ID,A.类型,A.名称,B.内容 From zlProcedure A,zlProceduretext B Where A.ID = B.过程ID And A.ID=[1] And B.性质 = 4"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", lngKey)
                If rs.BOF = False Then
                    '得到过程内容并执行此过程
                    strTmp = NVL(rs("内容").value)
                    strTmp = "CREATE OR REPLACE " & Trim(strTmp) & vbCrLf & "/"
                End If
                If strTmp <> "" Then
                    Call gcnOracle.Execute(strTmp)
                    strSQL = "Zl_Zlprocedure_Delete(" & lngKey & ")"
                    Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                    Call SQLRecordExecute(rsSQL)
                End If
            End With
        '--------------------------------------------------------------------------------------------------------------
        Case "删除过程"
            
            '查询是否有过程在调用当前过程
            
            strSQL = "Select ID,名称 From zlProcedure Where ID=[1]"
            Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", vsf(0).RowData(vsf(0).Row))
            If rs.BOF = False Then
                strTmp = NVL(rs("名称").value)
                strSQL = "Select Distinct Name As 过程名称 from (Select Distinct Name,Type,Text From User_Source Where type in ('PROCEDURE','FUNCTION') and upper(Text) Like [1] And Name <> [2])"
                Set rs = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "", "%" & UCase(strTmp) & "%", UCase(strTmp))
                If rs.BOF = False Then
                    If mfrmProcedureRelating Is Nothing Then
                        Set mfrmProcedureRelating = New frmProcRelating
                    End If
                    Call mfrmProcedureRelating.ShowDialog(Me, vsf(0).RowData(vsf(0).Row), rs)
                Else
                    If MsgBox("确定删除过程“" & strTmp & "”吗?", vbInformation + vbOKCancel, "中联软件") = vbOK Then
                        strSQL = "Zl_Zlprocedure_Delete(" & vsf(0).RowData(vsf(0).Row) & ")"
                        Call gclsBase.SQLRecordAdd(rsSQL, strSQL)
                        Call SQLRecordExecute(rsSQL)
                    Else
                        Exit Function
                    End If
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case "移除过程"
            If vsf(0).Rows > 2 Then
                vsf(0).RemoveItem vsf(0).Row
                mclsVsf.AppendRows = True
            Else
                Call mclsVsf(0).ClearGrid
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case "连接数据库"
'            Call gclsPackage.OraDataOpen("192.168.4.19", "ZLHIS", "ZLHIS", "HIS")
        End Select
    Next
    ExecuteCommand = True
    Exit Function
errHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngLoop As Long
    Dim objControl As CommandBarControl
    Dim strSQL As String
    
    
    On Error GoTo errHand
    
    Dim rs As ADODB.Recordset
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_NewItem
        If mfrmProgramEdit Is Nothing Then
            Set mfrmProgramEdit = New frmProcEdit
        End If
        Call mfrmProgramEdit.ShowMe(Me, 0)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Modify
        If mfrmProgramEdit Is Nothing Then Set mfrmProgramEdit = New frmProcEdit
        
        If vsf(0).RowData(vsf(0).Row) > 0 Then
            If mfrmProgramEdit.ShowMe(Me, vsf(0).RowData(vsf(0).Row)) Then
                Call ExecuteCommand("刷新数据")
                Call ExecuteCommand("刷新状态")
            End If
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Disuse
    
        If MsgBox("您确认升级完成了吗？" & vbCrLf & "此操作会将本次升级前的过程记录更为上次过程记录！", vbOKCancel + vbInformation + vbDefaultButton2, "中联软件") = vbOK Then
            
            gcnOracle.Execute "Zl_Zlproceduretext_Move()"
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Audit
        If mfrmDifferenceCheck Is Nothing Then
            Set mfrmDifferenceCheck = New frmProcDiffrentCheck
        End If
        If mfrmDifferenceCheck.ShowMe(Me) Then
            Call ExecuteCommand("刷新数据")
            Call ExecuteCommand("刷新状态")
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Manage_Change_PaitNote
        Set rs = gclsBase.GetProcByState(1, 2)
        If rs.BOF = False Then
            MsgBox "检测到有过程还未调整完成，请先进行调整后再生成。", vbInformation + vbOKOnly, "中联软件"
            Exit Sub
        End If
        If mfrmBuildScript Is Nothing Then
            Set mfrmBuildScript = New frmProcBuildScript
        End If
        Call mfrmBuildScript.ShowMe(Me)
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        If ExecuteCommand("删除过程") Then
            Call ExecuteCommand("移除过程")
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Untread
        If MsgBox("确定要将此过程恢复为之前备份的标准过程吗？", vbOKCancel + vbInformation, "中联软件") = vbOK Then
            If ExecuteCommand("恢复过程") Then
                Call ExecuteCommand("移除过程")
            End If
            
        End If
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Word
        If mfrmCollectUpdate Is Nothing Then
            Set mfrmCollectUpdate = New frmProcCollectUpdate
        End If
        If mfrmCollectUpdate.ShowMe(Me) Then
            Call ExecuteCommand("刷新数据")
            Call ExecuteCommand("刷新状态")
        End If
    Case conMenu_File_Exit
        Unload Me
        
        
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_ToolBar_Button     '工具栏
    
        For lngLoop = 2 To cbsMain.Count
            cbsMain(lngLoop).Visible = Not cbsMain(lngLoop).Visible
        Next
        cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Text      '按钮文字
    
        For lngLoop = 2 To cbsMain.Count
            For Each objControl In cbsMain(lngLoop).Controls
                If objControl.Type = xtpControlButton Then
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
        Next
        cbsMain.RecalcLayout
        
    Case conMenu_View_ToolBar_Size      '大图标
    
        cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
        cbsMain.RecalcLayout
        
    Case conMenu_View_StatusBar         '状态栏
    
        stbThis.Visible = Not stbThis.Visible
        cbsMain.RecalcLayout
    
    Case conMenu_Help_Help              '帮助主题
    
'        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((ParamInfo.系统号) / 100))
        
    Case conMenu_Help_Web_Home 'Web上的中联
    
        Call zlHomePage(Me.hwnd)
        
    Case conMenu_Help_Web_Forum '中联论坛
    
        Call zlWebForum(Me.hwnd)
        
    Case conMenu_Help_Web_Mail '发送反馈
    
        Call zlMailTo(Me.hwnd)
        
    Case conMenu_Help_About '关于
        
        Call ShowAbout(Me)
        
    End Select
    Exit Sub
errHand:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    With vsf(0)
    Select Case Control.Id
    Case conMenu_Edit_Delete
        Control.Visible = (opt(2).value = True)
    Case conMenu_Edit_Untread
        Control.Visible = (opt(0).value = True Or opt(1).value = True)

    Case conMenu_View_ToolBar_Button            '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text              '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size              '大图标
        Control.Checked = cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar                 '状态栏
        Control.Checked = stbThis.Visible
    End Select

    End With
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case 1
        Item.Handle = picPane(0).hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    Call ExecuteCommand("初始数据")
    Call ExecuteCommand("刷新数据")
    Call ExecuteCommand("刷新状态")
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    gstrSplite = "FRCHENUSERPROCEDUREFRCHEN"
    Call ExecuteCommand("初始控件")
'    Call ExecuteCommand("连接数据库")
End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'    picPane(1).Move 0, -15, Me.ScaleWidth, picPane(1).ScaleHeight + 30
'    mclsVsf.AppendRows = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsVsf Is Nothing) Then
        Set mclsVsf = Nothing
    End If
    If Not (mclsVsfMark Is Nothing) Then
        Set mclsVsfMark = Nothing
    End If
    If Not (mfrmBuildScript Is Nothing) Then
        Unload mfrmBuildScript
    End If
    
    If Not (mfrmCollectUpdate Is Nothing) Then
        Unload mfrmCollectUpdate
    End If
    
    If Not (mfrmDifferenceCheck Is Nothing) Then
        Unload mfrmDifferenceCheck
    End If
    
    If Not (mfrmProcedureRelating Is Nothing) Then
        Unload mfrmProcedureRelating
    End If
    
    If Not (mfrmProgramEdit Is Nothing) Then
        Unload mfrmProgramEdit
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mintProcType = (Index + 1)
    Call ExecuteCommand("刷新数据")
    Call ExecuteCommand("刷新状态")
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        
        picPane(2).Move 15, 15, picPane(Index).ScaleWidth - 30
        vsf(0).Move 15, picPane(2).Top + picPane(2).Height + 15, picPane(Index).ScaleWidth - 30, picPane(Index).ScaleHeight - (picPane(2).Top + picPane(2).Height + 15) - 15
        mclsVsf.AppendRows = True
    Case 2
        txtLocation.Move picPane(Index).ScaleWidth - txtLocation.Width - 75
        lbl1.Move txtLocation.Left - lbl1.Width - 30
    End Select
    
End Sub

Private Sub txtLocation_GotFocus()
    Call gclsBase.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim intCol As Integer
    
    If KeyAscii = vbKeyReturn Then
        intCol = vsf(0).ColIndex("过程")
        lngRow = mclsVsf.FindRow(UCase(txtLocation.Text), intCol, 2, vsf(0).Row + 1)
        If lngRow = -1 Then
            lngRow = mclsVsf.FindRow(UCase(txtLocation.Text), intCol, 2)
        End If
        If lngRow > 0 And vsf(0).Row <> lngRow Then
            vsf(0).Row = lngRow
            vsf(0).ShowCell vsf(0).Row, vsf(0).Col
        End If
        
        Call gclsBase.LocationObj(txtLocation)
    End If
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Select Case Index
    Case 1
        If OldRow = NewRow Then Exit Sub
    End Select
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cbrPopupBar As CommandBar
    
    Select Case Button
    '------------------------------------------------------------------------------------------------------------------
    Case 2          '弹出菜单处理
        Call gclsBase.SendLMouseButton(vsf(Index).hwnd, X, Y)
        Select Case Index
        Case 0
            Set cbrPopupBar = gclsBase.CopyMenu(cbsMain, 2)
            If cbrPopupBar Is Nothing Then Exit Sub
            cbrPopupBar.ShowPopup
        End Select
        
    End Select
End Sub

