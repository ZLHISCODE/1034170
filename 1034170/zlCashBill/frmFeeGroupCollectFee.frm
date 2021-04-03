VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFeeGroupCollectFee 
   BorderStyle     =   0  'None
   Caption         =   "财务组收款管理"
   ClientHeight    =   7440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picCurrentMoney 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   345
      ScaleWidth      =   5865
      TabIndex        =   6
      Top             =   480
      Width           =   5895
      Begin VB.Label lblCurrentMoney 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "当前暂存金:    现金:3000元    支票:5000元    医保基金:10000元    个人账户:100元"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   8295
      End
   End
   Begin VB.PictureBox picSubWorker 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   3000
      ScaleHeight     =   4215
      ScaleWidth      =   1935
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
      Begin MSComctlLib.ListView lvwSubWorker 
         Height          =   4335
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   7646
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilsWorker"
         SmallIcons      =   "ilsWorkerSmall"
         ColHdrIcons     =   "ilsWorkerSmall"
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "姓名"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "编号"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "简码"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "所属部门"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.PictureBox picGeneralInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   5760
      ScaleHeight     =   2655
      ScaleWidth      =   3735
      TabIndex        =   0
      Top             =   4440
      Width           =   3735
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   8
         Top             =   450
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFeeGroupCollectFee.frx":0000
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VB.TextBox txtSendNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   8
         Width           =   2500
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCollectorInfo 
         Height          =   1095
         Left            =   0
         TabIndex        =   3
         Top             =   420
         Width           =   2655
         _cx             =   4683
         _cy             =   1931
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         GridColor       =   -2147483636
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmFeeGroupCollectFee.frx":054E
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
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VB.Label lblInfo 
         Caption         =   "轧帐单号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   840
      End
   End
   Begin MSComctlLib.ImageList ilsWorker 
      Left            =   480
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":073F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":1019
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsWorkerSmall 
      Left            =   1200
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":18F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupCollectFee.frx":1E8D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   240
      Top             =   3720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpCollectFees 
      Bindings        =   "frmFeeGroupCollectFee.frx":2427
      Left            =   480
      Top             =   600
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFeeGroupCollectFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjChargeBillCollect As New clsChargeBill, mfrmChargeBillTotalCollect As Form  '收款信息和票据对象
Private mlngModule As Long, mstrPrivs As String
Private mlngGroupID As Long '缴款组ID
Private mfrmMain As Form    '主窗体
Private mcbrListView As CommandBar, mcbrControl As CommandBarControl

Private Enum EM_Pan
    EM_Pan_人员表 = 1
    EM_Pan_收费轧帐信息 = 2
    EM_Pan_收款及票据信息 = 3
    EM_Pan_人员余额 = 4
End Enum

Private Sub SetGrid()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSF控件
    '编制:刘尔旋
    '日期:2013-10-13
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    
    With vsCollectorInfo
        For i = 0 To .Cols - 1
            If .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            If .ColKey(i) = "ID" Or .ColKey(i) = "过滤" Or .ColKey(i) = "收款部门" Or .ColKey(i) = "选择" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Then .ColData(i) = "1|0"
        Next
    End With
    
    zl_vsGrid_Para_Restore mlngModule, vsCollectorInfo, Me.Caption, "收费员轧帐信息", False

End Sub

Public Sub ClearChargeAndBillTotalForm()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:外部调用清除票据窗体内容
    '编制:刘尔旋
    '日期:2013-10-12
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Call mobjChargeBillCollect.ClearChargeAndBillTotalForm
End Sub

Public Sub ChargeRollingListShow(frmMain As Object, bytType As TotalType, strIDs As String)
    Call mobjChargeBillCollect.ChargeRollingListShow(frmMain, bytType, strIDs, mlngModule, mstrPrivs)
End Sub

Public Sub InitMe(frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngGroupID As Long)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:初始化收款界面
    '入参:frmMain-主窗体
    '     lngModule-模块号
    '     strPrivs-权限串
    '     lngGroupID-组ID
    '编制:刘尔旋
    '日期:2013-10-10
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain
    mlngModule = lngModule
    mlngGroupID = lngGroupID
    mstrPrivs = strPrivs
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '返回:设置成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-09-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
        
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
    '初始化设置
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With

    cbsThis.EnableCustomization False
    
    Set mcbrListView = cbsThis.Add("人员表弹出菜单", xtpBarPopup)
    With mcbrListView.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标")
        mcbrControl.Checked = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细列表")
    End With
    
    cbsThis.ActiveMenuBar.Visible = False
    
    zlDefCommandBars = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetListViewMenu() As CommandBar
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:获取弹出菜单
    '返回:弹出菜单
    '编制:刘尔旋
    '日期:2013-10-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Set GetListViewMenu = mcbrListView
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrMenuView As CommandBarPopup
    Set cbrMenuView = frmFeeGroupManage.GetListViewMenu
    Select Case Control.ID
        Case conMenu_View_LargeICO
            mcbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = True
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = True
            mcbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            Call ChangeListViewType(1)
        Case conMenu_View_MinICO
            mcbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_MinICO).Checked = True
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = True
            mcbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            Call ChangeListViewType(2)
        Case conMenu_View_ListICO
            mcbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_ListICO).Checked = True
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = True
            mcbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = False
            Call ChangeListViewType(3)
        Case conMenu_View_DetailsICO
            mcbrListView.Controls.Find(, conMenu_View_LargeICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_LargeICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_MinICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_MinICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_ListICO).Checked = False
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_ListICO).Checked = False
            mcbrListView.Controls.Find(, conMenu_View_DetailsICO).Checked = True
            cbrMenuView.CommandBar.Controls.Find(, conMenu_View_DetailsICO).Checked = True
            Call ChangeListViewType(4)
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmChargeBillTotalCollect Is Nothing Then Unload mfrmChargeBillTotalCollect
    Set mobjChargeBillCollect = Nothing
End Sub

Private Sub lvwSubWorker_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mcbrListView.ShowPopup
    End If
End Sub

Private Sub picCurrentMoney_Resize()
    On Error Resume Next
    With lblCurrentMoney(0)
        .Top = 15
        .Width = picCurrentMoney.Width - 15
        .Height = picCurrentMoney.Height - 15
    End With
End Sub

Private Sub txtSendNO_GotFocus()
    Call SelAll(txtSendNO)
End Sub

Private Sub dkpCollectFees_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionAttaching Then Cancel = True
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub txtSendNO_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandle
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 65 Or KeyAscii > 90) And _
       (KeyAscii < 97 Or KeyAscii > 122) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        If txtSendNO.Text = "" Then
            KeyAscii = 0
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Dim i As Integer, strSQL As String
        Dim rsTmp As New ADODB.Recordset
        '完全匹配输入单号
        With vsCollectorInfo
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("轧帐单号")) = txtSendNO.Text Then
                    If .Enabled And .Visible Then .SetFocus
                    DoEvents
                    .Select i, .ColIndex("选择")
                    .TopRow = i
                    Exit Sub
                End If
            Next i
            strSQL = "Select 收款员" & vbNewLine & _
                     "From 人员收缴记录" & vbNewLine & _
                     "Where 记录性质 = 1 And 缴款组id = [1] And (小组收款人 = [3] Or 小组收款人 Is Null) And 作废时间 Is Null And 小组收款id Is Null And NO = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, txtSendNO.Text, UserInfo.姓名)
            If rsTmp.RecordCount <> 0 Then
                LoadWorkerCollectDetail (Nvl(rsTmp!收款员))
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("轧帐单号")) = txtSendNO.Text Then
                        If .Enabled And .Visible Then .SetFocus
                        DoEvents
                        .Select i, 1
                        .TopRow = i
                        Exit Sub
                    End If
                Next i
            End If
        End With
        
        '自动调整输入单号,再次进行查找
        txtSendNO.Text = GetFullNO(txtSendNO.Text, 137)
        With vsCollectorInfo
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("轧帐单号")) = txtSendNO.Text Then
                    If .Enabled And .Visible Then .SetFocus
                    DoEvents
                    .Select i, .ColIndex("选择")
                    .TopRow = i
                    Exit Sub
                End If
            Next i
            strSQL = "Select 收款员" & vbNewLine & _
                     "From 人员收缴记录" & vbNewLine & _
                     "Where 记录性质 = 1 And 缴款组id = [1] And 作废时间 Is Null And (小组收款人 = [3] Or 小组收款人 Is Null) And 小组收款id Is Null And NO = [2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, txtSendNO.Text, UserInfo.姓名)
            If rsTmp.RecordCount <> 0 Then
                LoadWorkerCollectDetail (Nvl(rsTmp!收款员))
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("轧帐单号")) = txtSendNO.Text Then
                        If .Enabled And .Visible Then .SetFocus
                        DoEvents
                        .Select i, 1
                        .TopRow = i
                        Exit Sub
                    End If
                Next i
            End If
        End With
        MsgBox "没有找到轧帐单号[" & txtSendNO.Text & "]的记录！", vbInformation, gstrSysName
        If txtSendNO.Visible Then txtSendNO.SetFocus
        Call SelAll(txtSendNO)
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub SetDockingPanel()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:创建DOCKINGPANEL控件
    '编制:刘尔旋
    '日期:2013-09-04
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim objPanel As Pane
    On Error GoTo errHandle
    With dkpCollectFees
        .SetCommandBars cbsThis
        .VisualTheme = ThemeOffice2003
        Set objPanel = .CreatePane(EM_Pan_人员表, 200, 1800, DockLeftOf)
        objPanel.Handle = picSubWorker.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.MinTrackSize.Width = 75
        objPanel.MaxTrackSize.Width = 150
        Set objPanel = .CreatePane(EM_Pan_收费轧帐信息, 2000, 800, DockRightOf, objPanel)
        objPanel.Handle = picGeneralInfo.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
        objPanel.Title = "收费轧账信息"
        objPanel.MinTrackSize.Height = 100
        Set objPanel = .CreatePane(EM_Pan_收款及票据信息, 2000, 1000, DockBottomOf, objPanel)
        objPanel.Handle = mfrmChargeBillTotalCollect.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.Title = "收款及票据信息"
        objPanel.MinTrackSize.Height = 230
        Set objPanel = .CreatePane(EM_Pan_人员余额, 2000, 100, DockBottomOf)
        objPanel.Handle = picCurrentMoney.hWnd
        objPanel.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
        objPanel.Title = "人员余额"
        objPanel.MinTrackSize.Height = 35
        objPanel.MaxTrackSize.Height = 35
        Set .PaintManager.CaptionFont = lblCurrentMoney(0).Font
        .Options.HideClient = True
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picSubWorker_Resize()
    On Error Resume Next
    lvwSubWorker.Width = picSubWorker.Width
    lvwSubWorker.Height = picSubWorker.Height
End Sub

Public Sub ChangeListViewType(ByVal intType As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:调整人员列表显示方式
    '入参:intType-列表显示方式: 1-大图标;2-小图标;3-列表;4-详细列表
    '编制:刘尔旋
    '日期:2013-10-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Select Case intType
        Case 1
            lvwSubWorker.View = lvwIcon
        Case 2
            lvwSubWorker.View = lvwSmallIcon
        Case 3
            lvwSubWorker.View = lvwList
        Case 4
            lvwSubWorker.View = lvwReport
    End Select
End Sub

Private Sub picGeneralInfo_Resize()
    On Error Resume Next
    With vsCollectorInfo
        .Width = picGeneralInfo.Width - 15
        .Height = picGeneralInfo.Height - 430
    End With
End Sub

Private Sub lvwSubWorker_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    lvwSubWorker.Drag 0
End Sub

Private Sub lvwSubWorker_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LoadWorkerCollectDetail(Item.Text)
End Sub

Private Sub lvwSubWorker_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '禁止编辑或者移动人员列表
    If Button = 1 Then
        If lvwSubWorker.HitTest(x, y) Is Nothing Then Exit Sub
        lvwSubWorker.Drag 1
    End If
End Sub

Public Sub AfterCollectEdit()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:小组收款完毕后刷新界面数据
    '编制:刘尔旋
    '日期:2013-09-12
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Call LoadWorkerCollectDetail(lvwSubWorker.SelectedItem.Text)
End Sub

Private Sub LoadWorkerCollectDetail(ByVal strWorker As String)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取指定的收费员的收费信息
    '入参:strWorker--收费员
    '编制:刘尔旋
    '日期:2013-09-09
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim strSQL As String, rsTmp As New ADODB.Recordset, i As Integer
    strSQL = "" & _
    "Select a.ID, a.NO, a.登记时间,decode(nvl(A.类别,0),0,'所有类别',1,'收费',2,'预交',21,'门诊预交',22,'住院预交',3,'结帐',4,'挂号',5,'就诊卡',6,'消费卡','-') as 轧帐类别, " & _
    "       a.开始时间, a.终止时间, a.冲预交款, a.借入合计, a.借出合计, a.摘要, a.收款员" & vbNewLine & _
    "From 人员收缴记录 A" & vbNewLine & _
    "Where a.记录性质 = 1 And a.缴款组id = [1] And (a.小组收款人 = [3] Or a.小组收款人 Is Null) And a.作废时间 Is Null And a.小组收款id Is Null And a.财务收款时间 Is Null And a.收款员 = [2]" & vbNewLine & _
    "Order by 登记时间 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID, strWorker, UserInfo.姓名)
    
    With vsCollectorInfo
        .Rows = 1
        If rsTmp.RecordCount <> 0 Then
            Do While Not rsTmp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("选择")) = 0
                '0-所有类别(按全额轧帐),1-收费,2-预交,3-结帐,4-挂号,5-就诊卡,6-消费卡
                .TextMatrix(.Rows - 1, .ColIndex("轧帐类别")) = Nvl(rsTmp!轧帐类别)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = Nvl(rsTmp!No)
                .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = Nvl(rsTmp!登记时间)
                .TextMatrix(.Rows - 1, .ColIndex("收款员")) = Nvl(rsTmp!收款员)
                '.TextMatrix(.Rows - 1, .ColIndex("收款部门")) = Nvl(rsTmp!部门名称)
                .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = Nvl(rsTmp!开始时间)
                .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = Nvl(rsTmp!终止时间)
                .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = Format(Nvl(rsTmp!冲预交款), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = Format(Nvl(rsTmp!借入合计), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = Format(Nvl(rsTmp!借出合计), "0.00")
                .TextMatrix(.Rows - 1, .ColIndex("备注")) = Nvl(rsTmp!摘要)
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = Nvl(rsTmp!ID)
                rsTmp.MoveNext
            Loop
            .AutoSize 1, .Cols - 1
            zl_vsGrid_Para_Restore mlngModule, vsCollectorInfo, Me.Caption, "收费员轧帐信息", False
            .ColWidth(.ColIndex("选择")) = 615
            .ColHidden(.ColIndex("选择")) = False
        End If
        If .Rows = 1 Then .Rows = 2
    End With
    
    Call RefreshCurrentMoney(0)
    mobjChargeBillCollect.ClearChargeAndBillTotalForm
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsCollectorInfo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    With vsCollectorInfo
        'If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then .Select 0, 0
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        mobjChargeBillCollect.LoadChargeAndBillTotalData Me, mlngModule, mstrPrivs, EM_收费员轧帐, .TextMatrix(.RowSel, .ColIndex("ID"))
        Call zl_VsGridRowChange(vsCollectorInfo, OldRow, NewRow, OldCol, NewCol)
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub vsCollectorInfo_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call zl_vsGrid_Para_Save(mlngModule, vsCollectorInfo, Me.Caption, "收费员轧帐信息", False)
End Sub

Private Sub vsCollectorInfo_DblClick()
    With vsCollectorInfo
        If .RowSel < 1 Or .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then Exit Sub
        Call ChargeRollingListShow(mfrmMain, EM_收费员轧帐, Val(.TextMatrix(.RowSel, .ColIndex("ID"))))
    End With
End Sub

Private Sub vsCollectorInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsCollectorInfo_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCollectorInfo
        If Col <> .ColIndex("选择") Then Cancel = True
        If .TextMatrix(.RowSel, .ColIndex("ID")) = "" Then
            Cancel = True
            Exit Sub
        End If
        .Select Row, .ColIndex("选择")
    End With
End Sub

Private Sub vsCollectorInfo_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsCollectorInfo.ColIndex("选择") Then Cancel = True
    If Col = 0 Then Cancel = True
End Sub

Private Sub vsCollectorInfo_GotFocus()
    With vsCollectorInfo
        If Val(.TextMatrix(1, .ColIndex("ID"))) <> 0 Then
            .Select 1, .ColIndex("选择")
        End If
        Call zl_VsGridGotFocus(vsCollectorInfo)
        .Cell(flexcpBackColor, 0, 0, 0, .Cols - 1) = .BackColorFixed
    End With
End Sub

Private Sub RefreshCurrentMoney(ByVal intPanel As Integer)
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:刷新界面暂存金
    '入参:intPanel-TAB界面序号
    '编制:刘尔旋
    '日期:2013-09-18
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    On Error GoTo errHandle
    strSQL = "Select 结算方式,余额 From 人员缴款余额 Where 收款员=[1] And 性质=1"
    If intPanel = 1 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名)
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lvwSubWorker.SelectedItem.Text)
    End If
    
    lblCurrentMoney(intPanel).Caption = " 当前暂存金:   "
    If rsTmp.RecordCount <> 0 Then
        Do While Not rsTmp.EOF
            If Val(Nvl(rsTmp!余额)) <> 0 Then
                lblCurrentMoney(intPanel).Caption = lblCurrentMoney(intPanel).Caption & rsTmp!结算方式 & ":" & rsTmp!余额 & "元   "
            End If
            rsTmp.MoveNext
        Loop
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function LoadSubWorkers() As Boolean
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取缴款组下属人员
    '出参:mlngGroupID-缴款组ID
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2013-09-03
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim lvwItem As ListItem
    strSQL = "Select 组名称,负责人ID From 财务缴款分组 Where (删除日期 Is Null or 删除日期 Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    
    If rsTmp.RecordCount = 0 Then
        LoadSubWorkers = False
        Exit Function
    End If
    
    dkpCollectFees.Panes(1).Title = Nvl(rsTmp!组名称)

    strSQL = "Select b.Id, b.编号, b.姓名, b.性别, b.简码, d.名称" & vbNewLine & _
             "From 缴款成员组成 A, 人员表 B, 部门人员 C, 部门表 D" & vbNewLine & _
             "Where a.成员id = b.Id And 组id = [1] And a.成员id = c.人员id And c.部门id = d.Id And c.缺省 = 1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    
    Do While Not rsTmp.EOF
        If rsTmp!性别 Like "*女*" Then
            Set lvwItem = lvwSubWorker.ListItems.Add(, "_" & rsTmp!ID, Nvl(rsTmp!姓名), 2, 2)
            lvwItem.SubItems(1) = Nvl(rsTmp!编号)
            lvwItem.SubItems(2) = Nvl(rsTmp!简码)
            lvwItem.SubItems(3) = Nvl(rsTmp!名称)
        Else
            '男或者性别不明的情况
            Set lvwItem = lvwSubWorker.ListItems.Add(, "_" & rsTmp!ID, Nvl(rsTmp!姓名), 1, 1)
            lvwItem.SubItems(1) = Nvl(rsTmp!编号)
            lvwItem.SubItems(2) = Nvl(rsTmp!简码)
            lvwItem.SubItems(3) = Nvl(rsTmp!名称)
        End If
        rsTmp.MoveNext
    Loop
    LoadSubWorkers = True
    Exit Function
errHandle:
    LoadSubWorkers = False
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub Form_Load()
    mobjChargeBillCollect.SetFontSize lblCurrentMoney(0).Font.Size
    Set mfrmChargeBillTotalCollect = mobjChargeBillCollect.GetChargeAndBillTotalForm
    Call zlDefCommandBars
    Call SetDockingPanel
    If LoadSubWorkers = False Then
        Call frmFeeGroupManage.FailInit
        Exit Sub
    End If
    Call SetGrid
    vsCollectorInfo.Select 0, 0
End Sub

Private Sub vsCollectorInfo_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsCollectorInfo)
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCollectorInfo, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsCollectorInfo, Me.Caption, "收费员轧帐信息", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
