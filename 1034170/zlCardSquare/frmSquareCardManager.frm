VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareCardManager 
   Caption         =   "消费卡管理"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   Icon            =   "frmSquareCardManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   12105
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   -360
      ScaleHeight     =   2565
      ScaleWidth      =   12405
      TabIndex        =   3
      Top             =   645
      Width           =   12405
      Begin VB.PictureBox picModify 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   4845
         ScaleHeight     =   465
         ScaleWidth      =   5850
         TabIndex        =   11
         Top             =   -90
         Visible         =   0   'False
         Width           =   5850
         Begin VB.CommandButton cmdSel 
            Caption         =   "…"
            Height          =   300
            Left            =   4080
            TabIndex        =   17
            Top             =   135
            Width           =   315
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "完成修改(&O)"
            Height          =   350
            Left            =   4770
            TabIndex        =   14
            Top             =   105
            Width           =   1230
         End
         Begin VB.CheckBox chk修改 
            Caption         =   "修改卡类型(&X)"
            Height          =   350
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Value           =   2  'Grayed
            Width           =   1875
         End
         Begin VB.TextBox txtEdit 
            Height          =   330
            Left            =   2115
            TabIndex        =   12
            Top             =   120
            Width           =   2280
         End
         Begin MSComCtl2.DTPicker dtp卡有效日期 
            Height          =   300
            Left            =   2115
            TabIndex        =   16
            Top             =   120
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   113115139
            CurrentDate     =   40156.0854282407
         End
         Begin VB.ComboBox cbo卡类型 
            Height          =   300
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   130
            Width           =   2310
         End
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "回收"
         Height          =   405
         Index           =   3
         Left            =   4050
         TabIndex        =   10
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "退卡"
         Height          =   405
         Index           =   2
         Left            =   3210
         TabIndex        =   9
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "失效卡"
         Height          =   405
         Index           =   1
         Left            =   2235
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.CheckBox chkStatus 
         Caption         =   "有效卡"
         Height          =   405
         Index           =   0
         Left            =   1305
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   105
         TabIndex        =   4
         Top             =   435
         Width           =   6825
         _cx             =   12039
         _cy             =   3625
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
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   24
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSquareCardManager.frx":6852
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   5
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmSquareCardManager.frx":6B7E
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblTag 
         AutoSize        =   -1  'True
         Caption         =   "当前卡信息"
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   4590
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   0
      Top             =   5175
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   8025
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSquareCardManager.frx":70CC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12356
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   645
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":7960
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareCardManager.frx":7CB4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   495
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSquareCardManager.frx":8008
      Left            =   1005
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSquareCardManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mblnFirst  As Boolean, mstrPrivs As String, mstrTitle As String    '功能标题
Private mlngModule As Long, mstrKey As String
Private Type Ty_CurrCardStartus '当前卡状态
    blnHaveData As Boolean
    bln回收 As Boolean
    bln退卡 As Boolean
    bln到期 As Boolean
    bln用户使用有效卡 As Boolean
    bln卡已消费 As Boolean
    bln允许充值回退 As Boolean
    bln停用卡 As Boolean
End Type
Private mTy_CardStartus As Ty_CurrCardStartus
Private Enum mPgIndex
    Pg_充值记录 = 250101
    Pg_回收记录 = 250102
    Pg_消费记录 = 250103
End Enum
Private Enum mPaneID
    Pane_Search = 1     '搜索条件
    Pane_CardLists = 2  '卡列表
    Pane_CardDetails = 3    '详细列表
End Enum
Private mlng接口编号 As Long
Private mrs消费卡接口 As ADODB.Recordset
Private mfrmSquareCardCallBack As frmSquareCardCallBack
Private WithEvents mfrmSquareCardConsume As frmSquareCardConsume
Attribute mfrmSquareCardConsume.VB_VarHelpID = -1
Private WithEvents mfrmSquareCardInFull As frmSquareCardInFul
Attribute mfrmSquareCardInFull.VB_VarHelpID = -1
Private WithEvents mfrmFilter As frmSquareCardFilter
Attribute mfrmFilter.VB_VarHelpID = -1
Private mPanSearch As Pane
Private mobjSubFrm As Collection
Private mfrmActive As Form
Private mArrFilter As Variant
Private Const mconMenu_Lable = 3999
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-11-20 16:02:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsCardList
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("卡号")) = "1|0"
        .ColData(.ColIndex("标志")) = "-1|1"
        .ColData(.ColIndex("当前余额")) = "1|0"
        If .ColIndex("ID") >= 0 Then
            .ColData(.ColIndex("ID")) = "-1|1"
            .ColHidden(.ColIndex("ID")) = True
        End If
    End With
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-11-19 15:15:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:
    Set mobjSubFrm = New Collection
    Set mfrmSquareCardInFull = New frmSquareCardInFul
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_充值记录, "充值信息", mfrmSquareCardInFull.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_充值记录
    mobjSubFrm.Add mfrmSquareCardInFull, CStr(objItem.Tag)
    '有充值或回退权限时，才显示出来
    '106681:李南春，2017/3/10，回退权限的全名是"回退充值"
    If zlCheckPrivs(mstrPrivs, "充值") Or zlCheckPrivs(mstrPrivs, "回退充值") Then
        objItem.Visible = True: i = 0
    Else
        objItem.Visible = False: i = 1:
    End If
    
    Set mfrmSquareCardCallBack = New frmSquareCardCallBack
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_回收记录, "回收信息", mfrmSquareCardCallBack.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_回收记录
    mobjSubFrm.Add mfrmSquareCardCallBack, CStr(objItem.Tag)


    Set mfrmSquareCardConsume = New frmSquareCardConsume
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_消费记录, "消费信息", mfrmSquareCardConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_消费记录
    mobjSubFrm.Add mfrmSquareCardConsume, CStr(objItem.Tag)

     With tbPage
        tbPage.Item(i).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区域
    '编制:刘兴洪
    '日期:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    If mfrmFilter Is Nothing Then Set mfrmFilter = New frmSquareCardFilter
    Call mfrmFilter.Init条件

    With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "条件设置": mPanSearch.Options = PaneNoCloseable
        mPanSearch.MinTrackSize.Width = 220: mPanSearch.MaxTrackSize.Width = 300
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    
    zlRestoreDockPanceToReg Me, dkpMan, "区域"

End Function
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否有数据
    '返回:当前控件有数据,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 14:24:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String
     
    zlIsHaveData = False
    If Me.ActiveControl Is vsCardList Then
        zlIsHaveData = mTy_CardStartus.blnHaveData
    Else
        'dd
    End If
End Function

Private Sub chkStatus_Click(Index As Integer)
    Call SetCardRowColHide
End Sub

Private Sub chk修改_Click()
    Call SetModifyEnabled
End Sub
Private Sub SetModifyEnabled()
    Dim blnEnabled As Boolean
    
    blnEnabled = chk修改.value = 1
    cmdSel.Visible = False
    With vsCardList
        cbo卡类型.Visible = False
        dtp卡有效日期.Visible = False
        txtEdit.Visible = False
        chk修改.Visible = True
        cmdModify.Visible = blnEnabled
        Select Case .Col
        Case .ColIndex("有效期")
            dtp卡有效日期.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("卡类型")
            cbo卡类型.Visible = blnEnabled
            picModify.Visible = True
        Case .ColIndex("限制类别")
            txtEdit.Visible = blnEnabled
            picModify.Visible = True
            cmdSel.Visible = blnEnabled
        Case Else
            picModify.Visible = False
        End Select
        picModify.Visible = zlCheckPrivs(mstrPrivs, "修改卡信息") And picModify.Visible
    End With
End Sub
Private Sub SetModiyCaption()
    With vsCardList
        Select Case .Col
        Case .ColIndex("有效期")
            chk修改.Caption = "修改“有效期”"
        Case .ColIndex("卡类型")
            chk修改.Caption = "修改“卡类型”"
        Case .ColIndex("限制类别")
            chk修改.Caption = "修改“限制类别”"
        Case Else
            chk修改.Visible = False
        End Select
    End With
End Sub
Private Sub SetModifyDefaultValue()
    Dim i As Long
    
    With vsCardList
        Select Case .Col
        Case .ColIndex("有效期")
            If .Row > 0 Then
                If Trim(.TextMatrix(.Row, .Col)) = "" Then
                     dtp卡有效日期.value = Null
                Else
                    If CDate(.TextMatrix(.Row, .Col)) < dtp卡有效日期.MinDate Then
                        dtp卡有效日期.value = dtp卡有效日期.MinDate
                    Else
                        dtp卡有效日期.value = CDate(.TextMatrix(.Row, .Col))
                    End If
                End If
            End If
            cmdSel.Visible = False
        Case .ColIndex("卡类型")
            If .Row > 0 Then
                For i = 0 To cbo卡类型.ListCount - 1
                    If InStr(1, cbo卡类型.List(i), Trim(.TextMatrix(.Row, .Col))) > 0 Then
                        cbo卡类型.ListIndex = i: Exit For
                    End If
                Next
            End If
            cmdSel.Visible = False
        Case .ColIndex("限制类别")
            If .Row > 0 Then
                txtEdit.Text = Trim(.TextMatrix(.Row, .Col))
                txtEdit.Tag = txtEdit.Text
            End If
            cmdSel = True
        Case Else
            chk修改.Visible = False
            cmdSel.Visible = False
        End Select
    End With

End Sub

Private Sub cmdModify_Click()
   Call SaveBatchUpdateCardInfor
End Sub

Private Sub cmdSel_Click()
    If cmdSel.Visible = False Then Exit Sub
    If Select收费类别选择器(txtEdit, "") = False Then Exit Sub
    zlCtlSetFocus txtEdit
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '搜索条件窗体
        Item.Handle = mfrmFilter.hWnd
    Case mPaneID.Pane_CardDetails   '详细卡信息
        Item.Handle = picList.hWnd
    Case mPaneID.Pane_CardLists '卡列表
        Item.Handle = picCardList.hWnd
    End Select
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开指定报表
    '入参:lngSys-系统号
    '     strReportCode报表编号
    '编制:刘兴洪
    '日期:2009-11-19 14:15:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng消费卡ID As Long, str卡号 As String, str卡类型 As String
    Dim str发卡人 As String, str发卡日期 As String
    With vsCardList
        If .Row < 0 Then Exit Sub
        lng消费卡ID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
        str卡号 = Trim(.TextMatrix(.Row, .ColIndex("卡号")))
        str卡类型 = Trim(.TextMatrix(.Row, .ColIndex("卡类型")))
        str发卡人 = Trim(.TextMatrix(.Row, .ColIndex("发卡人")))
        str发卡日期 = Trim(.TextMatrix(.Row, .ColIndex("发卡时间")))
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "消费卡ID=" & lng消费卡ID, "卡号=" & str卡号, "卡类型=" & str卡类型, "发卡人=" & str发卡人, "发卡日期=" & str发卡日期)
End Sub

Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化菜单及工具栏
    '入参:
    '出参:
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-18 16:53:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
        
      
    Err = 0: On Error GoTo Errhand:
    '-----------------------------------------------------
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
    '-----------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "收费轧帐(&M)")
        mcbrControl.IconId = 227
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "重打缴款单(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "参数设置"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): mcbrControl.BeginGroup = True
    End With


    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "修改(&M)"):
         'Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBatchModify, "批量修改(&L)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡(&B)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "取消退卡(&K)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "回收(&H)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "取消回收(&S)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "卡片启用(&F)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "卡片停用(&P)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "充值(&C)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "充值回退(&T)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "修改密码(&G)"): mcbrControl.BeginGroup = True
        mcbrControl.Enabled = zlCheckPrivs(mstrPrivs, "修改密码")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "修改密码")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord_Force, "强制修改密码(&O)")
        mcbrControl.Enabled = zlCheckPrivs(mstrPrivs, "强制修改密码")
        mcbrControl.Visible = zlCheckPrivs(mstrPrivs, "强制修改密码")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): mcbrControl.BeginGroup = True
    End With
    
    '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("L"), conMenu_Edit_CardBathPay
        
        .Add FCONTROL, Asc("M"), conMenu_Edit_CardModify
        .Add FCONTROL, Asc("C"), conMenu_Edit_CardInFullBack
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_Edit_RollingCurtain
    End With
    
    '设置不常用菜单
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "发卡"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardModify, "修改")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "退卡"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelBack, "取消退卡")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCallBack, "回收"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelCallBack, "取消回收")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardResume, "卡片启用")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardStop, "卡片停用"): mcbrControl.BeginGroup = True
                
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "充值"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "充值作废")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "收费轧帐(&M)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    
    Set mcbrComboxToolBar = cbsThis.Add("消费卡接口", xtpBarTop)
    mcbrComboxToolBar.ShowTextBelowIcons = False
    mcbrComboxToolBar.ContextMenuPresent = False
    mcbrComboxToolBar.EnableDocking xtpFlagStretched
    
    With mcbrComboxToolBar.Controls
        Set mcbrControl = .Add(xtpControlLabel, mconMenu_Lable, "消费卡接口")
        'objControl.Flags = xtpFlagRightAlign
        Set objComBar = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "消费卡接口")
        'objComBar.Flags = xtpFlagRightAlign
        objComBar.Flags = xtpFlagControlStretched
        Dim intIndex As Integer
        intIndex = 1
        With mrs消费卡接口
            Do While Not .EOF
                objComBar.AddItem Nvl(!编号) & "-" & Nvl(!名称)
                objComBar.ItemData(intIndex) = Val(Nvl(!编号))
                If mlng接口编号 = Val(Nvl(!编号)) Then
                   objComBar.ListIndex = intIndex
                End If
                intIndex = intIndex + 1
                .MoveNext
            Loop
        End With
        If intIndex > 1 And objComBar.ListIndex <= 0 Then
            objComBar.ListIndex = 1:
        End If
        If objComBar.ListIndex > 0 Then
             mlng接口编号 = objComBar.ItemData(objComBar.ListIndex)
        End If
        
        objComBar.Width = 120:
        
    End With
 
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub PrintReBill()
    '功能:重打票据
    Dim lngID As Long, lng发卡序号 As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    
    On Error GoTo errHandle
    
    strTemp = zlCommFun.ShowMsgbox("缴款单打印", "请选择你要打印的缴款单", "发卡(&F),充值(&I),取消(&C)", Me, vbDefaultButton2)
    If strTemp = "取消" Or strTemp = "" Then Exit Sub
    
    If strTemp = "发卡" Then
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID <= 0 Then
                ShowMsgbox "没选中相关的消费卡"
                Exit Sub
            End If
            
        End With
        gstrSQL = "Select 发卡序号 From 消费卡目录 where id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
        lng发卡序号 = Val(Nvl(rsTemp!发卡序号))
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "付款序号=" & lng发卡序号, "缴款=" & 0, "找补=" & 0, "充值ID=0", "ReportFormat=1", 2)
    Else
        lngID = mfrmSquareCardInFull.zlGet充值ID
        If lngID <= 0 Then
            ShowMsgbox "未选中相关的充值记录"
            Exit Sub
        End If
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, "充值ID=" & lngID, "缴款=" & 0, "找补=" & 0, "付款序号=0", "ReportFormat=2", 2)
    End If
    

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long
    Dim ctrCombox As CommandBarComboBox
    '------------------------------------
        
    Select Case Control.id
    'bytMode=1 打印;2 预览;3 输出到EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill       '"重打缴款单(&R)")
        Call PrintReBill
    Case conMenu_Edit_RollingCurtain   '收费轧帐
          Call zlExecuteChargeRollingCurtain(Me)
    Case conMenu_Edit_CardPay    '发卡(&S)")
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_发卡, mlng接口编号) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
        
    Case conMenu_Edit_CardBathPay    '批量发卡(&P)")
    Case conMenu_Edit_CardModify    '修改(&M)"):
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID < 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("当前状态")) <> "有效" And .TextMatrix(.Row, .ColIndex("当前状态")) <> "失效" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_修改, mlng接口编号, lngID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardBack    '退卡(&B)"): mcbrControl.BeginGroup = True
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("当前状态")) <> "有效" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_退卡, mlng接口编号, lngID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelBack   '取消退卡
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("当前状态")) <> "退卡" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_取消退卡, mlng接口编号, lngID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCallBack    '回收(&H)")
        With vsCardList
'            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
'            If lngID <= 0 Then Exit Sub
            'If .TextMatrix(.Row, .ColIndex("当前状态")) <> "有效" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_回收, mlng接口编号, 0) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardCancelCallBack  '取消回收
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID <= 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("当前状态")) <> "回收" Then Exit Sub
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_取消回收, mlng接口编号, lngID) = False Then Exit Sub
        If MsgBox("数据已经发生改变,是否重新刷新数据?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToRpt
    Case conMenu_Edit_CardResume        '卡片启用
        If SaveCardResumeAndStop(False) = False Then Exit Sub
    Case conMenu_Edit_CardStop        '卡片停用
        If SaveCardResumeAndStop(True) = False Then Exit Sub
    Case conMenu_Edit_CardInFull    '充值(&C)")
        With vsCardList
            lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
            If lngID <= 0 Then lngID = 0
            If .TextMatrix(.Row, .ColIndex("当前状态")) <> "有效" Then lngID = 0
            If mTy_CardStartus.bln用户使用有效卡 = False Then lngID = 0
        End With
        If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_充值, mlng接口编号, lngID) = False Then Exit Sub
        Call mfrmSquareCardInFull.zlReLoadData(mlng接口编号, lngID)
    Case conMenu_Edit_CardInFullBack    '充值回退(&T)")
         
        If mfrmSquareCardInFull.zl充值取消 = False Then Exit Sub
    Case conMenu_Edit_ChangePassWord    '修改密码
        If frmModiCardPass.zlModifyPass(Me, mlngModule, mlng接口编号, True) Then
            Exit Sub
        End If
    Case conMenu_Edit_ChangePassWord_Force  '强制修改密码
        If frmModiCardPass.zlModifyPass(Me, mlngModule, mlng接口编号, False) Then
            Exit Sub
        End If
    Case conMenu_COMBOX_INTERFACE   '点击选择
        Set ctrCombox = Control
        mlng接口编号 = ctrCombox.ItemData(ctrCombox.ListIndex)
        Call LoadDataToRpt
    Case conMenu_View_Refresh   '刷新
        '重新刷新数据
        Call LoadDataToRpt
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean
    If Me.Visible = False Then Exit Sub
    
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_Edit_RollingCurtain    '收费轧帐
        Control.Visible = zlCheckPrivs(mstrPrivs_RollingCurtain, "轧帐")
        Control.Enabled = Control.Visible
    Case conMenu_File_PrintSingleBill           '"重打缴款单(&R)"
        Control.Visible = zlCheckPrivs(mstrPrivs, "消费卡收费收据")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardPay, conMenu_Edit_CardBathPay   '发卡(&S),批量发卡(&P)
        Control.Visible = zlCheckPrivs(mstrPrivs, "发卡")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardModify   ', conMenu_Edit_CardBatchModify   '修改(&M),批量修改(&L)
        Control.Visible = zlCheckPrivs(mstrPrivs, "修改卡信息")
        Control.Enabled = Control.Visible And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardBack    '退卡(&B)
        Control.Visible = zlCheckPrivs(mstrPrivs, "退卡")
        Control.Enabled = Control.Visible And (Not mTy_CardStartus.bln卡已消费) And mTy_CardStartus.bln用户使用有效卡
    Case conMenu_Edit_CardCancelBack  '取消退卡
        Control.Visible = zlCheckPrivs(mstrPrivs, "退卡")
        Control.Enabled = Control.Visible And (Not mTy_CardStartus.bln卡已消费) And mTy_CardStartus.bln退卡
    Case conMenu_Edit_CardCallBack   '回收(&H),批量回收(&J)
        Control.Visible = zlCheckPrivs(mstrPrivs, "回收")
        Control.Enabled = Control.Visible
    
    Case conMenu_Edit_CardCancelCallBack  '取消回收
        Control.Visible = zlCheckPrivs(mstrPrivs, "回收")
        Control.Enabled = Control.Visible And mTy_CardStartus.bln回收
    Case conMenu_Edit_CardResume        '卡片启用
        Control.Visible = zlCheckPrivs(mstrPrivs, "卡片启用")
        Control.Enabled = Control.Visible And mTy_CardStartus.bln停用卡 And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardStop        '卡片停用
        Control.Visible = zlCheckPrivs(mstrPrivs, "卡片停用")
        Control.Enabled = Control.Visible And Not mTy_CardStartus.bln停用卡 And mTy_CardStartus.blnHaveData
    Case conMenu_Edit_CardInFull    '充值(&C)")
        Control.Visible = zlCheckPrivs(mstrPrivs, "充值")
        Control.Enabled = Control.Visible       ' And mTy_CardStartus.bln用户使用有效卡
        
    Case conMenu_Edit_CardInFullBack    '充值回退(&T)")
        Control.Visible = zlCheckPrivs(mstrPrivs, "回退充值")
        Control.Enabled = Control.Visible And mTy_CardStartus.bln允许充值回退
    Case conMenu_View_Refresh   '刷新
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_File_Parameter     '参数调用
            frmSquareCardParaSet.ShowParaSet Me, mlngModule, mstrPrivs
        Case Else   '其他操作功能调用
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    
    zl_CtlSetFocus vsCardList
    Call vsCardList_GotFocus
    mblnFirst = False
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String
    Dim i As Long
    mblnFirst = True
    
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mlng接口编号 = Val(zlDatabase.GetPara("上次接口号", glngSys, mlngModule, 0, , InStr(1, mstrPrivs, ";参数设置;") > 0))
    strShow = Trim(zlDatabase.GetPara("卡显示方式", glngSys, mlngModule, "1011", Array(chkStatus(0), chkStatus(1), chkStatus(2), chkStatus(3)), InStr(1, mstrPrivs, ";参数设置;") > 0))
    If Len(strShow) < 4 Then strShow = strShow & "11111"
    For i = 0 To 3
        chkStatus(i).value = IIf(Val(Mid(strShow, i + 1, 1)) = 1, 1, 0)
    Next
    dtp卡有效日期.MinDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    dtp卡有效日期.value = DateAdd("m", 1, dtp卡有效日期.MinDate)
    dtp卡有效日期.value = Null
    chk修改.value = 0
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitPanel
    Call InitPage
    Call zlDefCommandBars '初始菜单及工具栏
    Call InitVsGrid
    Set mArrFilter = mfrmFilter.GetFilterCon
    Call LoadDataToRpt
    '设置状态栏相关的颜色
    zlSetStatusPanelCololor Me, stbThis, 3, "停用", vbRed, False
    zlSetStatusPanelCololor Me, stbThis, 4, "回收", vbBlue, False
    zlSetStatusPanelCololor Me, stbThis, 5, "失效", &HFF00FF, False
    zlSetStatusPanelCololor Me, stbThis, 6, "有效", Me.ForeColor, False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long, strTemp As String
   SaveWinState Me, App.ProductName, mstrTitle
   
    strTemp = ""
    For i = 0 To 3
        strTemp = strTemp & IIf(chkStatus(i).value = 1, 1, 0)
    Next
   
   zlDatabase.SetPara "卡显示方式", strTemp, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
   zlDatabase.SetPara "上次接口号", mlng接口编号, glngSys, mlngModule, InStr(1, mstrPrivs, ";参数设置;") > 0
   
   zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlCheckPrivs(mstrPrivs, "参数设置")
   
   zlSaveDockPanceToReg Me, dkpMan, "区域"
   
    '关闭子窗口
    For i = 1 To mobjSubFrm.count
        If Not mobjSubFrm(i) Is Nothing Then Unload mobjSubFrm(i)
    Next
    If Not frmModiCardPass Is Nothing Then Unload frmModiCardPass
End Sub
Private Sub SetCardRowColHide(Optional lngLocalRow As Long = -1)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置行的显示和隐藏
    '入参:lngLocalRow -指定行(-1代表全部重新设置)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-22 21:05:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngRows As Long, i As Long
    Dim lngCurRow As Long
    
    Err = 0: On Error GoTo Errhand:
    
    With vsCardList
        i = 1: lngRows = .Rows - 1
        If lngLocalRow < 0 Then
            .Redraw = flexRDNone
        Else
            i = lngLocalRow: lngRows = lngLocalRow
        End If
        
        For lngRow = i To lngRows
            '1-有效, 2-回收,3-退卡,4-失效,8-停用
            .RowHidden(lngRow) = False
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("当前状态")))
            Case 1
                If chkStatus(0).value = 0 Then .RowHidden(lngRow) = True
            Case 2
                If chkStatus(3).value = 0 Then .RowHidden(lngRow) = True
            Case 3
                If chkStatus(2).value = 0 Then .RowHidden(lngRow) = True
            Case 4
                If chkStatus(1).value = 0 Then .RowHidden(lngRow) = True
            End Select
            If .RowHidden(lngRow) = False Then
                If lngCurRow < .Row Then lngCurRow = lngRow
            End If
        Next
        If lngLocalRow < 0 Then
            If lngCurRow > 0 And .RowHidden(.Row) = True Then .Row = lngCurRow
            .Redraw = flexRDBuffered
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume

End Sub


Private Function zlPopuMenus(ByVal blnListView As Boolean) As Boolean
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Err = 0: On Error Resume Next
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Function
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next

    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls

            Select Case mcbrControl.id
            Case conMenu_View_ShowStoped, conMenu_View_ShowAll, conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                cbrPopupItem.Checked = mcbrControl.Checked
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Function
Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查数据依赖性
    '返回:数据合法,返回true，否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    zlCheckDepend = False
 
    On Error GoTo errHandle
    
    gstrSQL = "Select 名称   From 结算方式 Where 性质 = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查现金结算方式", UserInfo.id)
    If rsTemp.EOF Then
        ShowMsgbox "结算方式中不存在一条件有现金性质的结算方式,请在结算方式管理中设置!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    Set mrs消费卡接口 = zlGet消费卡接口
    mrs消费卡接口.Filter = "自制卡=1"
    If mrs消费卡接口.RecordCount = 0 Then
        ShowMsgbox "消费卡接口中不存在相应的消费卡接口,不能进行维护!"
        Exit Function
    End If
    
    Set rsTemp = zlGet收费类别
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   没有相关的收费项目类别,请与系统管理员联系!"
        Exit Function
    End If
    gstrSQL = "Select rownum as ID, 编码,名称, 缺省面额, 缺省折扣, 缺省标志 From 消费卡类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "消费卡类型")
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   没有设置相关的消费卡类型,请在[字典管理]中设置!"
        Exit Function
    End If
    zlComboxLoadFromRecodeset Me.Caption, rsTemp, cbo卡类型, True
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,显示相关的项目及分类信息
    '编制:刘兴洪
    '日期:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not zlCheckDepend Then Exit Sub            '数据依赖性测试
    Me.Caption = strTitle
    RestoreWinState Me, App.ProductName, mstrTitle
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    
    vRect = GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlCheckPrivs(mstrPrivs, "参数设置")
    
End Sub

Private Sub mfrmFilter_zlRefreshCon(ByVal arrFilter As Variant)
    Set mArrFilter = arrFilter
    '重新加载数据
    Call LoadDataToRpt
End Sub
Private Function LoadDataToRpt() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-11-19 15:43:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, strSubWhere As String, lngRow As Long, lngPre消费ID As Long
    Dim rsTemp As ADODB.Recordset, strCurDate As String
    
    Err = 0: On Error GoTo Errhand:
    strSubWhere = ""
    
    If mArrFilter("发卡时间")(0) <> "1901-01-01" And mArrFilter("回收时间")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (发卡时间 Between [1] And [2] Or 回收时间 Between [3] And [4])"
    ElseIf mArrFilter("发卡时间")(0) = "1901-01-01" And mArrFilter("回收时间")(0) <> "1901-01-01" Then
        strSubWhere = strSubWhere & " And (回收时间 Between [3] And [4])"
    ElseIf mArrFilter("发卡时间")(0) <> "1901-01-01" And mArrFilter("回收时间")(0) = "1901-01-01" Then
        strSubWhere = strSubWhere & " And (发卡时间 Between [1] And [2])"
    End If
    If mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) <> "" Then
        strSubWhere = strSubWhere & " And (卡号 Between [5] And [6])"
    ElseIf mArrFilter("卡号范围")(0) = "" And mArrFilter("卡号范围")(1) <> "" Then
        strWhere = strWhere & " And A.卡号=[6]"
    ElseIf mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) = "" Then
        strWhere = strWhere & " And A.卡号=[5]"
    End If
    If strSubWhere = "" Then
        '如果没有结定时间范围,就只能查找当前的领卡人和发卡人
        If mArrFilter("领卡人") <> "" Then strWhere = strWhere & " and  A.领卡人 like [7]"
        If mArrFilter("发卡人") <> "" Then strWhere = strWhere & " and  A.发卡人 like [8]"
    Else
        If mArrFilter("领卡人") <> "" Then strSubWhere = strSubWhere & " and  领卡人 like [7]"
        If mArrFilter("发卡人") <> "" Then strSubWhere = strSubWhere & " and  发卡人 like [8]"
    End If
    
    If Trim(mArrFilter("卡类型")) <> "所有" Then strWhere = strWhere & " and  A.卡类型 = [9]"
    
    If Val(mArrFilter("包含停用卡")) = 1 Then
        strWhere = strWhere & " And  A.当前状态 <= 9"   '需要用到索引
    Else
        strWhere = strWhere & " And  A.当前状态+0 <= 9 And A.停用日期 >= To_Date('3000-01-01', 'yyyy-mm-dd')"   '需要用到索引
    End If
    
    If strSubWhere <> "" Then
        strSubWhere = Mid(Trim(strSubWhere), 4)
        gstrSQL = "" & _
        "   Select A.ID, A.卡类型, A.卡号, A.序号, A.密码, A.限制类别, A.可否充值, A.有效期, A.发卡原因, A.发卡人, A.领卡人, " & _
        "          A.发卡时间, A.回收人, A.回收时间, A.当前状态, A.备注, A.卡面金额, A.销售金额, A.充值折扣率, A.余额, A.停用人, " & _
        "          A.停用日期,decode(A1.编码,NULL,'',A1.编码||'-'||A1.名称) as 领卡部门,decode(C.消费卡id,NULL,0,1) as 消费" & _
        "   From 消费卡目录 A,部门表 A1, " & _
        "        (Select 接口编号,消费卡id From 病人卡结算记录  where 接口编号=[10] Group By 接口编号,消费卡id Having Count(*)>0) C, " & _
        "        (Select 卡号 ,max(序号) as 序号  From 消费卡目录  Where " & strSubWhere & "  Group by  卡号) B" & _
        "   Where  a.id=c.消费卡ID(+) and a.领卡部门id=A1.ID(+) And c.接口编号(+)=[10] And  A.卡号 = B.卡号 and a.序号=b.序号 and A.接口编号=[10]  " & strWhere
    Else
        gstrSQL = "" & _
        "   Select A.ID, A.卡类型, A.卡号, A.序号, A.密码, A.限制类别, A.可否充值, A.有效期, A.发卡原因, A.发卡人, A.领卡人, " & _
        "          A.发卡时间, A.回收人, A.回收时间, A.当前状态, A.备注, A.卡面金额, A.销售金额, A.充值折扣率, A.余额, A.停用人, " & _
        "          A.停用日期,decode(A1.编码,NULL,'',A1.编码||'-'||A1.名称) as 领卡部门,decode(C.消费卡id,NULL,0,1) as 消费" & _
        "   From 消费卡目录 A,部门表 A1, " & _
        "        (Select 接口编号,消费卡id From 病人卡结算记录  where 接口编号=[10] Group By 接口编号,消费卡id Having Count(*)>0) C, " & _
        "   Where  a.id=c.消费卡ID(+) and a.领卡部门id=A1.ID(+) and A.接口编号=[10]   " & strWhere
    End If
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        CDate(mArrFilter("发卡时间")(0)), CDate(mArrFilter("发卡时间")(1)), _
        CDate(mArrFilter("回收时间")(0)), CDate(mArrFilter("回收时间")(1)), _
        CStr(mArrFilter("卡号范围")(0)), CStr(mArrFilter("卡号范围")(1)), _
        CStr(mArrFilter("领卡人")), CStr(mArrFilter("发卡人")), _
        CStr(mArrFilter("卡类型")), mlng接口编号)
    
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    With vsCardList
        If .Row > 0 Then
            lngPre消费ID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
        End If
        .Redraw = flexRDNone
        .Clear 1
        .Rows = 2
        .Cell(flexcpForeColor, 1, .FixedCols, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("卡号")) = Nvl(rsTemp!卡号)
            .Cell(flexcpData, lngRow, .ColIndex("卡号")) = Nvl(rsTemp!id)
            
            .TextMatrix(lngRow, .ColIndex("卡类型")) = Nvl(rsTemp!卡类型)
            .Cell(flexcpData, .Row, .ColIndex("密码")) = Nvl(rsTemp!密码)
            .TextMatrix(lngRow, .ColIndex("密码")) = "********"
            
            .TextMatrix(lngRow, .ColIndex("充值卡")) = IIf(Val(Nvl(rsTemp!可否充值)) = 1, "√", "")
            
            
            .TextMatrix(lngRow, .ColIndex("有效期")) = Format(rsTemp!有效期, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("有效期"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("有效期")) = ""
            
            .TextMatrix(lngRow, .ColIndex("发卡人")) = Nvl(rsTemp!发卡人)
            .TextMatrix(lngRow, .ColIndex("发卡原因")) = Nvl(rsTemp!发卡原因)
            .TextMatrix(lngRow, .ColIndex("发卡时间")) = Format(rsTemp!发卡时间, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("领卡人")) = Nvl(rsTemp!领卡人)
            .TextMatrix(lngRow, .ColIndex("领卡部门")) = Nvl(rsTemp!领卡部门)
            
            
            .TextMatrix(lngRow, .ColIndex("停用人")) = Nvl(rsTemp!停用人)
            .TextMatrix(lngRow, .ColIndex("停用日期")) = Format(rsTemp!停用日期, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("停用日期"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("停用日期")) = ""
            
            .TextMatrix(lngRow, .ColIndex("回收人")) = Nvl(rsTemp!回收人)
            .TextMatrix(lngRow, .ColIndex("回收时间")) = Format(rsTemp!回收时间, "yyyy-mm-dd HH:MM:SS")
            If Trim(.TextMatrix(lngRow, .ColIndex("回收时间"))) >= "3000-01-01" Then .TextMatrix(lngRow, .ColIndex("回收时间")) = ""
            
            .TextMatrix(lngRow, .ColIndex("限制类别")) = Nvl(rsTemp!限制类别)
            .TextMatrix(lngRow, .ColIndex("充值折扣率")) = Format(rsTemp!充值折扣率, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("面额")) = Format(rsTemp!卡面金额, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("销售金额")) = Format(rsTemp!销售金额, "###0.00;-###0.00;;")
            
            .TextMatrix(lngRow, .ColIndex("当前余额")) = Format(rsTemp!余额, "###0.00;-###0.00;;")
            .TextMatrix(lngRow, .ColIndex("已消费")) = IIf(Val(Nvl(rsTemp!消费)) = 1, "√", "")
            .TextMatrix(lngRow, .ColIndex("备注")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("当前状态")) = ""
            '1-有效, 2-回收,3-退卡
            .Cell(flexcpData, .Row, .ColIndex("有效期")) = ""
            '1-有效, 2-回收,3-退卡,4-失效,8-停用
            If Format(rsTemp!有效期, "yyyy-mm-dd HH:MM:SS") <= strCurDate Then
                .TextMatrix(lngRow, .ColIndex("当前状态")) = "失效"
                .Cell(flexcpData, lngRow, .ColIndex("当前状态")) = 4
                .Cell(flexcpData, lngRow, .ColIndex("有效期")) = "4"   '失效了
            Else
                .TextMatrix(lngRow, .ColIndex("当前状态")) = Decode(Val(Nvl(rsTemp!当前状态)), 1, "有效", 2, "回收", 3, "退卡", 4, "失效", ""):
                .Cell(flexcpData, lngRow, .ColIndex("当前状态")) = Val(Nvl(rsTemp!当前状态)) Mod 10
            End If
            If lngPre消费ID = Val(Nvl(rsTemp!id)) Then
                .Row = lngRow
                If .RowIsVisible(.Row) = False Then .TopRow = .Row
            End If
            '设置颜色行
            Call SetGridRowForeColor(lngRow)
            SetCardRowColHide lngRow
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If .Row <= 0 Then .Row = 1
        .Redraw = flexRDBuffered
    End With
    zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "卡信息列表", True, True
    
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    LoadDataToRpt = True
    Exit Function
Errhand:
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据状态，设置行颜色
    '编制:刘兴洪
    '日期:2009-11-20 15:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long, int状态 As Integer
    With vsCardList
        If .TextMatrix(lngRow, .ColIndex("停用日期")) <> "" Then
            lngColor = vbRed
        ElseIf Val(.Cell(flexcpData, lngRow, .ColIndex("有效期"))) = 4 Then
            lngColor = &HFF00FF
        Else
            Select Case Val(.Cell(flexcpData, lngRow, .ColIndex("当前状态")))
            Case 2, 3
                  lngColor = vbBlue
            Case Else
                '1-有效, 2-回收,3-退卡,4-失效,8-停用
                lngColor = &H80000008
            End Select
        End If
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = lngColor
    End With
End Sub

Private Sub mfrmSquareCardConsume_zlDblClick(ByVal lng结算ID As Long, ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If InStr(1, mstrPrivs, ";卡结算消费明细帐;") = 0 Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_INSIDE_1503_1", Me, "卡结算ID=" & lng结算ID, 1)
End Sub
 

Private Sub mfrmSquareCardInFull_AfterRowChange(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    mTy_CardStartus.bln允许充值回退 = mfrmSquareCardInFull.zl允许回退
    
End Sub

Private Sub mfrmSquareCardInFull_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    '弹出菜单:充值相关
    Dim cbrPopupBar As CommandBar, cbrPopupItem As CommandBarControl
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    With cbrPopupBar.Controls
        If zlCheckPrivs(mstrPrivs, "充值") Then Set cbrPopupItem = .Add(xtpControlButton, conMenu_Edit_CardInFull, "充值(&C)")
        If zlCheckPrivs(mstrPrivs, "回退") Then Set cbrPopupItem = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "充值回退(&T)"): cbrPopupItem.BeginGroup = True
        Set cbrPopupItem = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
    End With
    cbrPopupBar.ShowPopup
End Sub

 
Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
        picModify.Width = .ScaleWidth - picModify.Left - 50
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub zlSetInitCardCustomType()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡列信息的当前状态
    '编制:刘兴洪
    '日期:2009-11-19 14:46:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    With mTy_CardStartus
        .blnHaveData = False: .bln到期 = False: .bln回收 = False
        .bln退卡 = False: .bln用户使用有效卡 = False: .bln卡已消费 = False
        .bln停用卡 = False
    End With
    
    With vsCardList
        If .Rows < 2 Then Exit Sub
        If .Row < 1 Then Exit Sub
        mTy_CardStartus.blnHaveData = Trim(.TextMatrix(1, .ColIndex("卡号"))) <> ""
        '1-有效, 2-回收,3-退卡,4-失效,8-停用
        mTy_CardStartus.bln回收 = Val(.Cell(flexcpData, .Row, .ColIndex("当前状态"))) = 2
        mTy_CardStartus.bln退卡 = Val(.Cell(flexcpData, .Row, .ColIndex("当前状态"))) = 3
        mTy_CardStartus.bln到期 = Val(.Cell(flexcpData, .Row, .ColIndex("有效期"))) = 1 '失效了的卡
        mTy_CardStartus.bln用户使用有效卡 = Val(.Cell(flexcpData, .Row, .ColIndex("当前状态"))) = 1   '表明有效使用的卡(在用户处使用的卡)
        mTy_CardStartus.bln卡已消费 = Val(.Cell(flexcpData, .Row, .ColIndex("已消费"))) = 1
        mTy_CardStartus.bln停用卡 = Trim(.TextMatrix(.Row, .ColIndex("停用日期"))) <> ""
    End With
    '下面检查充值回退:
    
End Sub

Private Sub picModify_Click()
    Err = 0: On Error Resume Next
    With picModify
        cmdModify.Left = .ScaleWidth - cmdModify.Width - 50
        txtEdit.Width = cmdModify.Left - txtEdit.Left
        cmdSel.Left = txtEdit.Left + txtEdit.Width - cmdSel.Width
        cbo卡类型.Width = txtEdit.Width
        dtp卡有效日期.Width = txtEdit.Width
    End With
End Sub

Private Sub txtEdit_Change()
    txtEdit.Tag = ""
End Sub

Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit.Text <> "" And txtEdit.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select收费类别选择器(txtEdit, Trim(txtEdit.Text)) = False Then Exit Sub
    zlCtlSetFocus txtEdit
End Sub

Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng消费卡ID As Long
    zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldCol <> NewCol Then
        Call SetModiyCaption: Call SetModifyEnabled
    End If
    
    If OldRow = NewRow Then Exit Sub
    zlCommFun.ShowFlash "正在装载数据,请稍候..."
    With vsCardList
        lng消费卡ID = Val(.Cell(flexcpData, NewRow, .ColIndex("卡号")))
        Call mfrmSquareCardCallBack.zlReLoadData(mlng接口编号, lng消费卡ID)  '回收记录
        Call mfrmSquareCardInFull.zlReLoadData(mlng接口编号, lng消费卡ID) '充值记录
        Call mfrmSquareCardConsume.zlReLoadData(mlng接口编号, lng消费卡ID)   '消费记录
    End With
    '设置行的信息
    Call zlSetInitCardCustomType
    Call SetModifyDefaultValue
    zlCommFun.StopFlash
End Sub

Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlCheckPrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "卡信息列表", True, zlCheckPrivs(mstrPrivs, "参数设置")
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim blnCardList As Boolean
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "消费卡清册"
    
    If CStr(mArrFilter("发卡时间")(0)) <> "1901-01-01" Then
        objRow.Add "发卡时间：" & CStr(mArrFilter("发卡时间")(0)) & "至" & CStr(mArrFilter("发卡时间")(1))
    End If
    If CStr(mArrFilter("回收时间")(0)) <> "1901-01-01" Then
        objRow.Add "回收时间：" & CStr(mArrFilter("回收时间")(0)) & "至" & CStr(mArrFilter("回收时间")(1))
    End If
    If objRow.count > 1 Then
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
    End If
    If mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) <> "" Then
        objRow.Add "卡号范围：" & CStr(mArrFilter("卡号范围")(0)) & "至" & CStr(mArrFilter("卡号范围")(1))
    ElseIf mArrFilter("卡号范围")(0) = "" And mArrFilter("卡号范围")(1) <> "" Then
        objRow.Add "卡号：" & CStr(mArrFilter("卡号范围")(1))
    ElseIf mArrFilter("卡号范围")(0) <> "" And mArrFilter("卡号范围")(1) = "" Then
        objRow.Add "卡号：" & CStr(mArrFilter("卡号范围")(0))
    End If
    
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    
    If mArrFilter("领卡人") <> "" Then objRow.Add "领卡人：" & mArrFilter("领卡人")
    If mArrFilter("发卡人") <> "" Then objRow.Add "发卡人：" & mArrFilter("发卡人")
    If mArrFilter("卡类型") <> "" Then objRow.Add "卡类型：" & mArrFilter("卡类型")
    If Val(mArrFilter("包含停用卡")) = 1 Then objRow.Add "包含停用卡"
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    '由于打印控件不能识别列隐藏属性
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
            
        Next
    End With
    
    Err = 0: On Error GoTo Errhand:
    Set objPrint.Body = vsCardList
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '恢复
    With vsCardList
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function SaveCardResumeAndStop(ByVal blnStop As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:卡片停用或启用
    '入参:blnStop-停用卡片
    '编制:刘兴洪
    '日期:2009-12-14 09:45:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, lngID As Long, lngRow As Long, i As Long
    With vsCardList
        lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
        If lngID <= 0 Then Exit Function
        If blnStop Then '停用
        
            If .TextMatrix(.Row, .ColIndex("停用日期")) <> "" Then Exit Function
            If MsgBox("你真的要对卡号为:“" & .TextMatrix(.Row, .ColIndex("卡号")) & "”的记录进行停用操作吗？" & vbCrLf & _
                        "   『是』: 进行停用操作,停用后的卡片将不能进行刷卡消费或不能再发卡！" & vbCrLf & _
                        "   『否』:放弃本次停用操作", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            If .TextMatrix(.Row, .ColIndex("停用日期")) = "" Then Exit Function
            If MsgBox("你真的要对卡号为:“" & .TextMatrix(.Row, .ColIndex("卡号")) & "”的记录进行启用操作吗？" & vbCrLf & _
                        "   『是』: 进行启用操作,启用后的卡片将能进行刷卡消费或回收回来的卡片能再发卡操作！" & vbCrLf & _
                        "   『否』:放弃本次启用用操作", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End With
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
   Err = 0: On Error GoTo Errhand:
    ' Zl_消费卡目录_Stopandresume
    gstrSQL = "Zl_消费卡目录_Stopandresume("
    '  Id_In       In 消费卡目录.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  停用人_In   In 消费卡目录.停用人%Type,
    gstrSQL = gstrSQL & IIf(blnStop = False, "NULL", "'" & UserInfo.姓名 & "'") & ","
    '  停用日期_In In 消费卡目录.停用日期%Type
    gstrSQL = gstrSQL & IIf(blnStop = False, "NULL", "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss')") & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With vsCardList
         If blnStop Then '停用
            If Val(mArrFilter("包含停用卡")) = 1 Then
                .TextMatrix(.Row, .ColIndex("停用日期")) = strDate
                .TextMatrix(.Row, .ColIndex("停用人")) = UserInfo.姓名
            Else
                lngRow = .Row
                If .Rows - 1 <= 2 Then
                      For i = 0 To .Cols - 1
                        .TextMatrix(.Rows - 1, i) = ""
                        .Cell(flexcpData, .Rows - 1, i) = ""
                      Next
                Else '删除行
                     .RemoveItem lngRow
                     If lngRow < .Rows - 1 Then
                        .Row = lngRow
                     Else
                        .Row = .Rows - 1
                     End If
                End If
            End If
        Else
            .TextMatrix(.Row, .ColIndex("停用日期")) = ""
            .TextMatrix(.Row, .ColIndex("停用人")) = ""
        End If
        Call SetGridRowForeColor(.Row)
    End With
    Call zlSetInitCardCustomType
    SaveCardResumeAndStop = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
 
Private Function SaveBatchUpdateCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:批量更新卡片信息
    '返回:更新成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 12:02:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFields As String, blnIsdate As Boolean, cllPro As New Collection, strFieldValue As String
    Dim strSQL As String, strIDIn As String, lngRow As Long
    
    With vsCardList
        Select Case .Col
        Case .ColIndex("有效期")
           strFields = "有效期": blnIsdate = True:
           If IsNull(dtp卡有效日期.value) Then
                strFieldValue = "3000-01-01 00:00:00"
           Else
                strFieldValue = Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM:SS")
           End If
        Case .ColIndex("卡类型")
           strFields = "卡类型": blnIsdate = False:
           If cbo卡类型.ListIndex < 0 Then
                ShowMsgbox "卡类型未选择,请选择卡类型"
                Exit Function
           End If
           strFieldValue = Mid(cbo卡类型.Text, InStr(1, cbo卡类型.Text, ".") + 1)
        Case .ColIndex("限制类别")
           strFields = "限制类别": blnIsdate = False:
           If txtEdit.Tag = "" And txtEdit.Text <> "" Then
                ShowMsgbox "限制类别选择错误,请检查!"
                Exit Function
           End If
           strFieldValue = Trim(txtEdit.Text)
        Case Else
           Exit Function
        End Select
        If MsgBox("你是否真的要批量修改“" & strFields & "”的值吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    End With
    Err = 0: On Error GoTo Errhand:
    strIDIn = ""
    With vsCardList
        For lngRow = 1 To .Rows - 1
            If zlCommFun.ActualLen(strIDIn) >= 3980 Then
                'Zl_消费卡目录_Batch_Update
                gstrSQL = "Zl_消费卡目录_Batch_Update("
                '  Ids_In    Varchar2,
                gstrSQL = gstrSQL & "'" & Mid(strIDIn, 2) & "',"
                '  字段_In   Varchar2,
                gstrSQL = gstrSQL & "'" & strFields & "',"
                '  字段值_In Varchar2,
                gstrSQL = gstrSQL & "'" & strFieldValue & "',"
                '  IsDate Number:=0
                gstrSQL = gstrSQL & " " & IIf(blnIsdate, 1, 0) & ")"
                AddArray cllPro, gstrSQL
                strIDIn = ""
            End If
            If Val(.Cell(flexcpData, lngRow, .ColIndex("卡号"))) <> 0 And .RowHidden(lngRow) = False Then
                strIDIn = strIDIn & "," & Val(.Cell(flexcpData, lngRow, .ColIndex("卡号")))
            End If
        Next
    End With
    If strIDIn <> "" Then
        'Zl_消费卡目录_Batch_Update
        gstrSQL = "Zl_消费卡目录_Batch_Update("
        '  Ids_In    Varchar2,
        gstrSQL = gstrSQL & "'" & Mid(strIDIn, 2) & "',"
        '  字段_In   Varchar2,
        gstrSQL = gstrSQL & "'" & strFields & "',"
        '  字段值_In Varchar2,
        gstrSQL = gstrSQL & "'" & strFieldValue & "',"
        '  IsDate Number:=0
        gstrSQL = gstrSQL & " " & IIf(blnIsdate, 1, 0) & ")"
        AddArray cllPro, gstrSQL
        strIDIn = ""
    End If
    If cllPro.count = 0 Then
        Exit Function
    End If
    ExecuteProcedureArrAy cllPro, Me.Caption
    
    '加载数据
    With vsCardList
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("卡号"))) <> 0 And .RowHidden(lngRow) = False Then
                Select Case .Col
                  Case .ColIndex("有效期")
                     If IsNull(dtp卡有效日期.value) Then
                          .TextMatrix(lngRow, .Col) = ""
                     Else
                          .TextMatrix(lngRow, .Col) = Format(dtp卡有效日期.value, "yyyy-mm-dd HH:MM:SS")
                     End If
                  Case .ColIndex("卡类型")
                       .TextMatrix(lngRow, .Col) = Mid(cbo卡类型.Text, InStr(1, cbo卡类型.Text, "-") + 1)
                  Case .ColIndex("限制类别")
                      .TextMatrix(lngRow, .Col) = Trim(txtEdit.Text)
                  Case Else
                     Exit Function
                  End Select
            End If
        Next
    End With
    SaveBatchUpdateCardInfor = True
    MsgBox "修改成功!"
    
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Function Select收费类别选择器(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能: 收费类别选择器
    '入参::objCtl-指定控件
    '     strSearch-要搜索的条件
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "收费类别选择器"
    vRect = GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSearch, False)
    
    gstrSQL = "" & _
        " Select rownum as ID,编码,名称,简码 From 收费项目类别"
    If strSearch <> "" Then
        gstrSQL = gstrSQL & _
        "           Where ( 编码 like upper([1]) or 名称 like [1] or 简码 like upper([1]) )"
    End If
    gstrSQL = gstrSQL & vbCrLf & " Order by 编码"
  
    Set rsTemp = frmItemSelectMulit.ShowSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, True, strKey)
 
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "没有满足条件的收费类别,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If objCtl.Enabled Then objCtl.SetFocus
    With rsTemp
        objCtl.Tag = ""
        Do While Not .EOF
            strTemp = strTemp & "," & Nvl(rsTemp!名称)
            objCtl.Tag = objCtl.Tag & "," & Nvl(rsTemp!名称)
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strKey = objCtl.Tag
    objCtl.Text = strTemp
    objCtl.Tag = strKey
    zlCommFun.PressKey vbKeyTab
    Select收费类别选择器 = True
End Function

Private Sub vsCardList_DblClick()
    Dim lngID As Long
    With vsCardList
        lngID = Val(.Cell(flexcpData, .Row, .ColIndex("卡号")))
        If lngID <= 0 Then Exit Sub
    End With
    If frmSquareSendCard.zlShowCard(Me, mlngModule, mstrPrivs, gEd_查询, mlng接口编号, lngID) = False Then Exit Sub
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLOSTFOCUS vsCardList, gSysColor.lngGridColorLost
End Sub
