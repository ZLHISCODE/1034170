VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.UserControl ClinicPlanUnit 
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ScaleHeight     =   9225
   ScaleWidth      =   12000
   Begin VB.PictureBox picPageSub 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   9420
      ScaleHeight     =   975
      ScaleWidth      =   1275
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1275
      Begin XtremeSuiteControls.TabControl tbPageSub 
         Height          =   930
         Index           =   1
         Left            =   60
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   30
         Width           =   1080
         _Version        =   589884
         _ExtentX        =   1905
         _ExtentY        =   1640
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPageSub 
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   8100
      ScaleHeight     =   975
      ScaleWidth      =   1275
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1275
      Begin XtremeSuiteControls.TabControl tbPageSub 
         Height          =   930
         Index           =   0
         Left            =   60
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   1080
         _Version        =   589884
         _ExtentX        =   1905
         _ExtentY        =   1640
         _StockProps     =   64
      End
   End
   Begin VB.CheckBox chkOnlyOneUse 
      Caption         =   "独占方式"
      Height          =   300
      Left            =   5310
      TabIndex        =   4
      Top             =   50
      Width           =   1035
   End
   Begin VB.PictureBox picFun 
      BorderStyle     =   0  'None
      Height          =   4065
      Left            =   7320
      ScaleHeight     =   4065
      ScaleWidth      =   765
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   765
      Begin VB.CommandButton cmdFun 
         Caption         =   "<<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   105
         TabIndex        =   11
         Top             =   1935
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   "<"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   1465
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">>"
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   995
         Width           =   555
      End
      Begin VB.CommandButton cmdFun 
         Caption         =   ">"
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   525
         Width           =   555
      End
   End
   Begin VB.PictureBox picUnit 
      BorderStyle     =   0  'None
      Height          =   4050
      Index           =   0
      Left            =   8130
      ScaleHeight     =   4050
      ScaleWidth      =   2760
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   2760
      Begin VB.CheckBox chkForbidBespeak 
         Caption         =   "禁止预约"
         Height          =   300
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   1110
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSelNum 
         Height          =   3285
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   360
         Width           =   2175
         _cx             =   3836
         _cy             =   5794
         Appearance      =   2
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
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ClinicPlanUnit.ctx":0000
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
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   930
      Left            =   8100
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   1080
      _Version        =   589884
      _ExtentX        =   1905
      _ExtentY        =   1640
      _StockProps     =   64
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按总量预约"
      Height          =   300
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Top             =   50
      Width           =   1200
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按比例预约"
      Height          =   300
      Index           =   0
      Left            =   1215
      TabIndex        =   1
      Top             =   50
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.OptionButton optBespeakMode 
      Caption         =   "按序号控制预约"
      Height          =   300
      Index           =   2
      Left            =   3690
      TabIndex        =   3
      Top             =   50
      Width           =   1560
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfNotSelNum 
      Height          =   4065
      Left            =   4890
      TabIndex        =   6
      Top             =   2400
      Width           =   2385
      _cx             =   4207
      _cy             =   7170
      Appearance      =   2
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":0070
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
   Begin VSFlex8Ctl.VSFlexGrid vsUnit 
      Height          =   2865
      Left            =   90
      TabIndex        =   5
      Top             =   405
      Width           =   3120
      _cx             =   5503
      _cy             =   5054
      Appearance      =   2
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanUnit.ctx":00E0
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
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "预约控制方式"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   110
      Width           =   1080
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H8000000A&
      Height          =   3615
      Left            =   1620
      Top             =   3810
      Width           =   2655
   End
End
Attribute VB_Name = "ClinicPlanUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mobj所有合作单位 As 合作单位控制集
Private mobj合作单位集 As 合作单位控制集
Private mobj所有号序集 As 号序信息集
Private mblnNotClick As Boolean
Private mblnEdit As Boolean
Private mblnValiedCanSave As Boolean

Private Enum PageSub_Index
    Pg_合作单位 = 0
    Pg_预约方式 = 1
End Enum

'属性变量:
Dim m_EditMode As gRegistPlanEditMode
Dim m_IsDataChanged As Boolean

'缺省属性值:
Const m_def_EditMode = 0
Const m_def_IsDataChanged = False
'事件声明:
Event DataIsChanged()


Public Function LoadData(ByVal obj合作单位集 As 合作单位控制集, ByVal obj所有号序集 As 号序信息集, _
    ByVal obj所有合作单位 As 合作单位控制集, Optional ByVal blnChanged As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载出诊安排
    '入参:
    '     obj合作单位集-合作单位分配信息
    '     obj所有合作单位 - 所有合作单位控制集 ,不传表示查看
    '     obj所有号序集 - 所有备选号序集
    '返回:加载成功，返回true,否则返回false
    '编制:刘兴洪
    '日期:2016-01-12 12:46:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobj合作单位集 = obj合作单位集
    Set mobj所有号序集 = obj所有号序集
    Set mobj所有合作单位 = obj所有合作单位

    If mobj合作单位集 Is Nothing Then Set mobj合作单位集 = New 合作单位控制集
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    m_IsDataChanged = blnChanged
    LoadData = InitData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitFace()
    Dim i As Integer
    
    Err = 0: On Error GoTo Errhand
    With tbPage.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .BoldSelected = True
        .Layout = xtpTabLayoutAutoSize
        .StaticFrame = False
        .ClientFrame = xtpTabFrameBorder
        .Position = xtpTabPositionBottom
    End With
    
    For i = tbPageSub.LBound To tbPageSub.UBound
        With tbPageSub(i).PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .Layout = xtpTabLayoutAutoSize
            .StaticFrame = False
            .ClientFrame = xtpTabFrameBorder
        End With
    Next
    
    tbPage.Enabled = False '禁用控件，否则从其他控件的按钮控件中切换到该控件时焦点并没有真正失去
    tbPage.InsertItem Pg_合作单位, "挂号合作单位", picPageSub(0).Hwnd, 0
    tbPage.InsertItem Pg_预约方式, "预约方式", picPageSub(1).Hwnd, 0
    tbPage.Item(Pg_合作单位).Selected = True
    tbPage.Enabled = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub UnitPageVisible(ByVal blnVisible As Boolean)
    '隐藏三方合作单位
    Dim i As Integer
    Dim blnDo As Boolean
    
    Err = 0: On Error GoTo Errhand
    'List
    For i = 1 To vsUnit.Rows - 1
        vsUnit.RowHidden(i) = False
        If vsUnit.RowData(i) = 1 Then vsUnit.RowHidden(i) = blnVisible = False
    Next
    'TabPage
    tbPageSub(Pg_合作单位).Visible = blnVisible
    If blnVisible = False Then
        tbPage.Enabled = False
        tbPage(Pg_预约方式).Selected = True
        tbPage.Enabled = True
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetGridColVisible(ByVal bln分时段 As Boolean, ByVal bln序号控制 As Boolean)
    '设置网格列可见状态
    Dim i As Integer, j As Integer
    
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.ColHidden(-1) = False
    vsfNotSelNum.AllowSelection = False
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).ColHidden(-1) = False
        vsfSelNum(i).Editable = flexEDNone '允许编辑
        vsfSelNum(i).FocusRect = flexFocusNone
        vsfSelNum(i).AllowSelection = False
    Next
    If bln分时段 Then
        If bln序号控制 Then
            '分时段序号控制"数量"列不可见
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("数量")) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("数量")) = True
                vsfSelNum(i).AllowSelection = True
            Next
        Else
            '分时段不序号控制"序号"列不可见
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("序号")) = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).Editable = flexEDKbdMouse  '允许编辑
                vsfSelNum(i).FocusRect = flexFocusLight
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("序号")) = True
            Next
        End If
    Else
        If bln序号控制 Then
            '不分时段序号控制只有"序号"列可见
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("时间段")) = True
            vsfNotSelNum.ColHidden(vsfNotSelNum.ColIndex("数量")) = True
            vsfNotSelNum.AllowSelection = True
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("时间段")) = True
                vsfSelNum(i).ColHidden(vsfSelNum(i).ColIndex("数量")) = True
                vsfSelNum(i).AllowSelection = True
            Next
        End If
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ClearUnitGridData()
    '清除表格数据
    Dim i As Integer
    
    With vsUnit
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                .TextMatrix(i, .ColIndex("禁止预约")) = 0
                .TextMatrix(i, .ColIndex("数量")) = ""
                .Cell(flexcpBackColor, i, .ColIndex("数量")) = vsUnit.BackColor
            End If
        Next
    End With
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '编制:刘兴洪
    '日期:2016-01-12 12:48:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim obj号序 As 号序信息, obj号序集 As 号序信息集
    Dim objVsfGrid As VSFlexGrid, obj合作单位 As 合作单位控制
    Dim bln分时段 As Boolean, bln序号控制 As Boolean, byt预约控制 As Byte
    Dim blnFind As Boolean, i As Long, lngRow As Long, j As Long
    
    Err = 0: On Error GoTo Errhand:
    
    '============================================
    '先加载所有合作单位，初始化网格
    picFun.Tag = ""
    If mobj所有合作单位 Is Nothing Then
        vsUnit.Clear 1: vsUnit.Rows = 1
    Else
        With vsUnit
            .Clear 1
            .Rows = 1
            .Rows = mobj所有合作单位.Count + 1
            lngRow = 1
            For Each obj合作单位 In mobj所有合作单位
                .TextMatrix(lngRow, .ColIndex("类型")) = IIf(obj合作单位.类型 = 1, "挂号合作单位", "预约方式")
                .TextMatrix(lngRow, .ColIndex("合作单位")) = obj合作单位.合作单位名称
                .RowData(lngRow) = obj合作单位.类型 '1-三方机构;2-预约方式
                lngRow = lngRow + 1
            Next
            .ColAlignment(.ColIndex("禁止预约")) = flexAlignCenterCenter
            .ColDataType(.ColIndex("禁止预约")) = flexDTBoolean
            
            '数据按类型分组
            .OutlineBar = flexOutlineBarComplete
            .Subtotal flexSTClear
            .Subtotal flexSTNone, .ColIndex("类型")
            .SubtotalPosition = flexSTAbove

            .Outline .ColIndex("合作单位")
            .OutlineCol = .ColIndex("合作单位")

            .MergeCompare = flexMCIncludeNulls
            .MergeCells = flexMergeRestrictRows
            For i = 0 To .Cols - 1
                .MergeCol(i) = True
            Next
            
            For i = 1 To .Rows - 1
                .MergeRow(i) = False
                If .IsSubtotal(i) Then
                    .IsCollapsed(i) = flexOutlineExpanded
                    .MergeRow(i) = True
                    .TextMatrix(i, .ColIndex("合作单位")) = .TextMatrix(i, .ColIndex("类型"))
                    .RowData(i) = IIf(.TextMatrix(i, .ColIndex("类型")) = "挂号合作单位", 1, 2)
                End If
            Next
        End With
    End If
    '加载页面
    Call InitUnitPage
    '清除表格数据
    Call ClearUnitGridData
    
    vsfNotSelNum.Clear 1: vsfNotSelNum.Rows = 1
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        vsfSelNum(i).Clear 1: vsfSelNum(i).Rows = 1
    Next
    '============================================
    
    bln分时段 = mobj所有号序集.是否分时段
    bln序号控制 = mobj所有号序集.是否序号控制
    '0-禁止预约(或挂号);1-按比例控制预约(或挂号);2-按总量控制预约(或挂号);3-按序号控制预约(或挂号);4-不作限制
    byt预约控制 = mobj合作单位集.预约控制方式
    
    '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    Call UnitPageVisible(mobj所有号序集.预约控制 <> 2)
    Call SetGridColVisible(bln分时段, bln序号控制)
    mblnEdit = bln分时段 And Not bln序号控制
    
    If bln分时段 = False And bln序号控制 = False And byt预约控制 = 3 Then byt预约控制 = 0
    mblnNotClick = True
    optBespeakMode(IIf(byt预约控制 = 0 Or byt预约控制 = 4, 0, byt预约控制 - 1)).Value = True
    chkOnlyOneUse.Value = IIf(mobj合作单位集.是否独占, vbChecked, vbUnchecked)
    mblnNotClick = False
    
    '标记按序号控制预约(或挂号)是否可见
    optBespeakMode(2).Tag = IIf(bln分时段 Or bln序号控制, "", "1")
    picFun.Tag = IIf(bln序号控制, "", "1")
    
    If byt预约控制 <> 3 Then
        For Each obj合作单位 In mobj合作单位集
            With vsUnit
                .Redraw = flexRDNone
                For i = 1 To .Rows - 1
                    If .RowData(i) = obj合作单位.类型 And .TextMatrix(i, .ColIndex("合作单位")) = obj合作单位.合作单位名称 Then
                        Select Case obj合作单位.预约控制方式
                        Case 0 '禁止预约
                            .TextMatrix(i, .ColIndex("禁止预约")) = 1
                            .Cell(flexcpBackColor, i, .ColIndex("数量")) = vbButtonFace
                        Case 1, 2
                            If Not obj合作单位.号序信息集 Is Nothing Then
                                For Each obj号序 In obj合作单位.号序信息集
                                    '序号:控制方式=0,1,2,4时，填为0;否则存储启用序号或分时段的序号
                                    '数量:控制方式=0,4时，填为0;控制方式=1时，存放比例,如20,代表20%,控制方式=2时，存储的是限约数量，比如：10表示只能预约10个号;控制方式=3时，存储限约数量，启用序号的，一般为1,不启用序号且分时段的，存储限约数量
                                    .TextMatrix(i, .ColIndex("数量")) = FormatEx(obj号序.数量, 2, False)
                                    Exit For
                                Next
                            End If
                        End Select
                    End If
                Next
                .Redraw = flexRDBuffered
            End With
        Next
    End If

    '加载所有序号信息
    If bln分时段 Or bln序号控制 Then
        With vsfNotSelNum
            .Redraw = flexRDNone
            For Each obj号序 In mobj所有号序集
                If obj号序.是否预约 And obj号序.数量 > 0 Then
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("序号")) = obj号序.序号
                    .TextMatrix(lngRow, .ColIndex("时间段")) = Format(obj号序.开始时间, "hh:mm") & "-" & Format(obj号序.终止时间, "hh:mm")
                    .Cell(flexcpData, lngRow, .ColIndex("时间段")) = obj号序.开始时间 & "-" & obj号序.终止时间
                    .TextMatrix(lngRow, .ColIndex("数量")) = obj号序.数量
                    .Cell(flexcpData, lngRow, .ColIndex("数量")) = obj号序.数量
                End If
            Next
            .Redraw = flexRDBuffered
        End With
        
        If bln分时段 And bln序号控制 = False Then
            For i = vsfSelNum.LBound To vsfSelNum.UBound
                With vsfSelNum(i)
                    .Redraw = flexRDNone
                    For Each obj号序 In mobj所有号序集
                        If obj号序.是否预约 And obj号序.数量 > 0 Then
                            .Rows = .Rows + 1
                            lngRow = .Rows - 1
                            .TextMatrix(lngRow, .ColIndex("序号")) = obj号序.序号
                            .TextMatrix(lngRow, .ColIndex("时间段")) = Format(obj号序.开始时间, "hh:mm") & "-" & Format(obj号序.终止时间, "hh:mm")
                            .Cell(flexcpData, lngRow, .ColIndex("时间段")) = obj号序.开始时间 & "-" & obj号序.终止时间
                            .TextMatrix(lngRow, .ColIndex("数量")) = 0
                        End If
                    Next
                    .Redraw = flexRDBuffered
                End With
            Next
        End If
        If vsfNotSelNum.Rows > 1 And vsfNotSelNum.Row < 1 Then vsfNotSelNum.Row = 1

        '加载合作单位已选择序号信息
        For Each obj合作单位 In mobj合作单位集
            Set objVsfGrid = GetUnitVsfGrid(obj合作单位.类型, obj合作单位.合作单位名称)
            If Not objVsfGrid Is Nothing Then
                Select Case obj合作单位.预约控制方式
                Case 0 '禁止预约
                    mblnNotClick = True
                    chkForbidBespeak(objVsfGrid.index).Value = vbChecked
                    mblnNotClick = False
                    objVsfGrid.Editable = flexEDNone
                Case 3
                    If Not obj合作单位.号序信息集 Is Nothing Then
                        vsfNotSelNum.Redraw = flexRDNone
                        objVsfGrid.Redraw = flexRDNone
                        For Each obj号序 In obj合作单位.号序信息集
                            '序号:控制方式=0,1,2,4时，填为0;否则存储启用序号或分时段的序号
                            '数量:控制方式=0,4时，填为0;控制方式=1时，存放比例,如20,代表20%,控制方式=2时，存储的是限约数量，比如：10表示只能预约10个号;控制方式=3时，存储限约数量，启用序号的，一般为1,不启用序号且分时段的，存储限约数量
                            If bln分时段 And bln序号控制 = False Then
                                RemoveItem vsfNotSelNum, objVsfGrid, obj号序.序号, True, obj号序.数量
                            Else
                                RemoveItem vsfNotSelNum, objVsfGrid, obj号序.序号
                            End If
                        Next
                        vsfNotSelNum.Redraw = flexRDBuffered
                        objVsfGrid.Redraw = flexRDBuffered
                    End If
                    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And mblnEdit, flexEDKbdMouse, flexEDNone)
                End Select
            End If
        Next
    End If
    
Handler:
    Call SetUnitVisible
    Call SetButtonEnable(SelectedVsfGridIndex)
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetUnitVsfGrid(ByVal byt类型 As Byte, ByVal str名称 As String) As VSFlexGrid
    '根据类型和名称获取对应的VSFlexGrid控件
    '入参：
    '   byt类型 1-三方机构;2-预约方式
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    If str名称 = "" Then Exit Function
     
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        If Val(vsfSelNum(i).Tag) = byt类型 And picUnit(i).Tag = str名称 Then
            Set GetUnitVsfGrid = vsfSelNum(i)
            Exit Function
        End If
    Next
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub chkForbidBespeak_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid, i As Long
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    Set objVsfGrid = vsfSelNum(index)
    
    objVsfGrid.Redraw = flexRDNone
    vsfNotSelNum.Redraw = flexRDNone
    If Not mobj所有号序集 Is Nothing Then
        If mobj所有号序集.是否分时段 And mobj所有号序集.是否序号控制 = False Then
            For i = 1 To objVsfGrid.Rows - 1
                RemoveItem vsfNotSelNum, objVsfGrid, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("序号"))), True, 0
            Next
            objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
            Exit Sub
        End If
    End If
    
    For i = 1 To objVsfGrid.Rows - 1
        If i > objVsfGrid.Rows - 1 Then Exit For
        RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("序号")))
        i = i - 1
    Next
    objVsfGrid.Redraw = flexRDBuffered
    vsfNotSelNum.Redraw = flexRDBuffered
    
    Call SetButtonEnable(objVsfGrid.index)
    objVsfGrid.Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(index).Value <> vbChecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetButtonEnable(ByVal index As Integer)
    If m_EditMode <> ED_RegistPlan_Edit _
        Or index < vsfSelNum.LBound Or index > vsfSelNum.UBound Then
        cmdFun(0).Enabled = False
        cmdFun(1).Enabled = False
        cmdFun(2).Enabled = False
        cmdFun(3).Enabled = False
        Exit Sub
    End If
    
    cmdFun(0).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Row > 0
    cmdFun(1).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfNotSelNum.Rows > 1
    cmdFun(2).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Row > 0
    cmdFun(3).Enabled = chkForbidBespeak(index).Value <> vbChecked And vsfSelNum(index).Rows > 1
End Sub

Private Sub RemoveItem(ByVal objVsfGridFrom As VSFlexGrid, ByVal objVsfGridTo As VSFlexGrid, ByVal lngSN As Long, _
    Optional ByVal blnChangeNum As Boolean, Optional lngNum As Long)
    '移动项目或更改数量
    '参数：
    '   lngSN 序号
    '   blnChangeNum 仅改变数量,分时段，不序号控制时
    '   lngNum 改变的数量
    Dim blnFind As Boolean, i As Integer, j As Integer
    Dim lngRow As Long
    Dim intLow As Integer, intHigh As Integer, intMid As Integer
    
    On Error GoTo Errhand
    If objVsfGridFrom.Rows > 1 Then
        If Val(objVsfGridFrom.TextMatrix(1, objVsfGridFrom.ColIndex("序号"))) = lngSN Then
            lngRow = 1
        ElseIf Val(objVsfGridFrom.TextMatrix(objVsfGridFrom.Rows - 1, objVsfGridFrom.ColIndex("序号"))) = lngSN Then
            lngRow = objVsfGridFrom.Rows - 1
        End If
    End If
    '二分法查找
    If lngRow = 0 Then
        intLow = 1
        intHigh = objVsfGridFrom.Rows - 1
        Do While intLow <= intHigh
            intMid = (intLow + intHigh) \ 2
            If Val(objVsfGridFrom.TextMatrix(intMid, objVsfGridFrom.ColIndex("序号"))) < lngSN Then '在后面
                intLow = intMid + 1
            ElseIf Val(objVsfGridFrom.TextMatrix(intMid, objVsfGridFrom.ColIndex("序号"))) > lngSN Then '在前面
                intHigh = intMid - 1
            Else
                lngRow = intMid: Exit Do
            End If
        Loop
    End If
    If lngRow = 0 Then Exit Sub
    
    If blnChangeNum Then
        For i = 1 To objVsfGridTo.Rows - 1
            If Val(objVsfGridTo.TextMatrix(i, objVsfGridTo.ColIndex("序号"))) = lngSN Then
                objVsfGridTo.TextMatrix(lngRow, objVsfGridTo.ColIndex("数量")) = lngNum
                Exit For
            End If
        Next
        '计算剩余数量
        lngNum = Val(objVsfGridFrom.Cell(flexcpData, lngRow, objVsfGridFrom.ColIndex("数量")))
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号"))) = lngSN Then
                    lngNum = lngNum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("数量")))
                    Exit For
                End If
            Next
        Next
        objVsfGridFrom.TextMatrix(lngRow, objVsfGridFrom.ColIndex("数量")) = lngNum
    Else
        '按顺序插入
        blnFind = False
        If objVsfGridTo.Rows <= 1 Then
            With objVsfGridFrom
                objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("序号")) & vbTab & .TextMatrix(lngRow, .ColIndex("时间段")) & _
                    vbTab & .TextMatrix(lngRow, .ColIndex("数量"))
                objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("时间段")) = .Cell(flexcpData, lngRow, .ColIndex("时间段"))
            End With
            blnFind = True
        Else
            If Val(objVsfGridTo.TextMatrix(1, objVsfGridTo.ColIndex("序号"))) >= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("序号")) & vbTab & .TextMatrix(lngRow, .ColIndex("时间段")) & _
                        vbTab & .TextMatrix(lngRow, .ColIndex("数量")), 1
                    objVsfGridTo.Cell(flexcpData, 1, objVsfGridTo.ColIndex("时间段")) = .Cell(flexcpData, lngRow, .ColIndex("时间段"))
                End With
                blnFind = True
            ElseIf Val(objVsfGridTo.TextMatrix(objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("序号"))) <= lngSN Then
                With objVsfGridFrom
                    objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("序号")) & vbTab & .TextMatrix(lngRow, .ColIndex("时间段")) & _
                        vbTab & .TextMatrix(lngRow, .ColIndex("数量"))
                    objVsfGridTo.Cell(flexcpData, objVsfGridTo.Rows - 1, objVsfGridTo.ColIndex("时间段")) = .Cell(flexcpData, lngRow, .ColIndex("时间段"))
                End With
                blnFind = True
            End If
        End If
        
        '二分法查找
        If blnFind = False Then
            intLow = 1
            intHigh = objVsfGridTo.Rows - 1
            Do While intLow <= intHigh
                intMid = (intLow + intHigh) \ 2
                If Val(objVsfGridTo.TextMatrix(intMid - 1, objVsfGridTo.ColIndex("序号"))) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("序号"))) > lngSN Then   '找到位置了，且肯定能找到
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("序号")) & vbTab & .TextMatrix(lngRow, .ColIndex("时间段")) & _
                            vbTab & .TextMatrix(lngRow, .ColIndex("数量")), intMid
                        objVsfGridTo.Cell(flexcpData, intMid, objVsfGridTo.ColIndex("时间段")) = .Cell(flexcpData, lngRow, .ColIndex("时间段"))
                    End With
                    Exit Do
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("序号"))) < lngSN _
                    And Val(objVsfGridTo.TextMatrix(intMid + 1, objVsfGridTo.ColIndex("序号"))) > lngSN Then '找到位置了，且肯定能找到
                    With objVsfGridFrom
                        objVsfGridTo.AddItem .TextMatrix(lngRow, .ColIndex("序号")) & vbTab & .TextMatrix(lngRow, .ColIndex("时间段")) & _
                            vbTab & .TextMatrix(lngRow, .ColIndex("数量")), intMid + 1
                        objVsfGridTo.Cell(flexcpData, intMid + 1, objVsfGridTo.ColIndex("时间段")) = .Cell(flexcpData, lngRow, .ColIndex("时间段"))
                    End With
                    Exit Do
                End If
                
                If Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("序号"))) < lngSN Then '在后面
                    intLow = intMid + 1
                ElseIf Val(objVsfGridTo.TextMatrix(intMid, objVsfGridTo.ColIndex("序号"))) > lngSN Then '在前面
                    intHigh = intMid - 1
                End If
            Loop
        End If
        objVsfGridFrom.RemoveItem lngRow
        
        If objVsfGridFrom.Rows > 1 And objVsfGridFrom.Row < 1 Then objVsfGridFrom.Row = 1
        If objVsfGridTo.Rows > 1 And objVsfGridTo.Row < 1 Then objVsfGridTo.Row = 1
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub chkForbidBespeak_GotFocus(index As Integer)
    chkForbidBespeak(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub chkForbidBespeak_LostFocus(index As Integer)
     chkForbidBespeak(index).BackColor = Me.BackColor
End Sub


Private Sub chkForbidBespeak_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkOnlyOneUse_Click()
    Dim i As Integer
    
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '清除表格数据
    Call ClearUnitGridData
End Sub

Private Sub chkOnlyOneUse_GotFocus()
    chkOnlyOneUse.BackColor = GCTRL_SELBACK_COLOR
End Sub
Private Sub chkOnlyOneUse_LostFocus()
     chkOnlyOneUse.BackColor = Me.BackColor
End Sub
Private Sub chkOnlyOneUse_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Property Get SelectedVsfGridIndex() As Integer
    '获取当前选择的网格索引
    Dim byt类型 As Byte, str名称 As String
    Dim objVsfGrid As VSFlexGrid
    
    On Error GoTo ErrHandler
    SelectedVsfGridIndex = -1
    
    With tbPageSub(tbPage.Selected.index)
        If .Selected Is Nothing Then Exit Sub
        str名称 = .Selected.Caption
        byt类型 = Val(.Selected.Tag)
    End With
     
    Set objVsfGrid = GetUnitVsfGrid(byt类型, str名称)
    If Not objVsfGrid Is Nothing Then
        SelectedVsfGridIndex = objVsfGrid.index
    End If
    Exit Property
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property

Private Sub cmdFun_Click(index As Integer)
    Dim objVsfGrid As VSFlexGrid, intSelectedGridIndex As Integer
    Dim blnFind As Boolean, i As Integer
    Dim intStartRow As Integer, intEndRow As Integer
    Dim byt类型 As Byte, str名称 As String
    
    On Error GoTo Errhand
    intSelectedGridIndex = SelectedVsfGridIndex
    If intSelectedGridIndex < vsfSelNum.LBound Or intSelectedGridIndex > vsfSelNum.UBound Then Exit Sub
    Set objVsfGrid = vsfSelNum(intSelectedGridIndex)
    
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    vsfNotSelNum.Redraw = flexRDNone
    objVsfGrid.Redraw = flexRDNone
    Select Case index
    Case 0 '选进
        '批量设置
        intStartRow = vsfNotSelNum.RowSel: intEndRow = vsfNotSelNum.Row
        If vsfNotSelNum.Row < vsfNotSelNum.RowSel Then
            intStartRow = vsfNotSelNum.Row: intEndRow = vsfNotSelNum.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(intStartRow, vsfNotSelNum.ColIndex("序号")))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < vsfNotSelNum.Rows Then vsfNotSelNum.Select intStartRow, 0
    Case 1 '全选进
        For i = 1 To vsfNotSelNum.Rows - 1
            If i > vsfNotSelNum.Rows - 1 Then Exit For
            RemoveItem vsfNotSelNum, objVsfGrid, Val(vsfNotSelNum.TextMatrix(i, vsfNotSelNum.ColIndex("序号")))
            i = i - 1
        Next
    Case 2 '移除
        '批量设置
        intStartRow = objVsfGrid.RowSel: intEndRow = objVsfGrid.Row
        If objVsfGrid.Row < objVsfGrid.RowSel Then
            intStartRow = objVsfGrid.Row: intEndRow = objVsfGrid.RowSel
        End If
        Do While True
            If intStartRow > intEndRow Then Exit Do
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(intStartRow, objVsfGrid.ColIndex("序号")))
            intEndRow = intEndRow - 1
        Loop
        If intStartRow > 0 And intStartRow < objVsfGrid.Rows Then objVsfGrid.Select intStartRow, 0
    Case 3 '全移除
        For i = 1 To objVsfGrid.Rows - 1
            If i > objVsfGrid.Rows - 1 Then Exit For
            RemoveItem objVsfGrid, vsfNotSelNum, Val(objVsfGrid.TextMatrix(i, objVsfGrid.ColIndex("序号")))
            i = i - 1
        Next
    End Select
    vsfNotSelNum.Redraw = flexRDBuffered
    objVsfGrid.Redraw = flexRDBuffered
    
    Call SetButtonEnable(intSelectedGridIndex)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdFun_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub optBespeakMode_Click(index As Integer)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mblnNotClick Then Exit Sub
    m_IsDataChanged = True: RaiseEvent DataIsChanged
    '清除表格数据
    Call ClearUnitGridData
    If Not mobj所有号序集 Is Nothing Then
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            chkForbidBespeak(i).Value = vbUnchecked
            For j = 1 To vsfSelNum(i).Rows - 1
                If mobj所有号序集.是否序号控制 Then
                    If j > vsfSelNum(i).Rows - 1 Then Exit For
                    RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号")))
                    j = j - 1
                Else
                    RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号"))), True, 0
                End If
            Next
        Next
    End If
    Call SetUnitVisible
    Call SetButtonEnable(SelectedVsfGridIndex)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub optBespeakMode_GotFocus(index As Integer)
    optBespeakMode(index).BackColor = GCTRL_SELBACK_COLOR
End Sub
 
Private Sub optBespeakMode_LostFocus(index As Integer)
     optBespeakMode(index).BackColor = Me.BackColor
End Sub

Private Sub optBespeakMode_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub


Private Sub picFun_Resize()
    Err = 0: On Error Resume Next
    cmdFun(0).Top = (picFun.ScaleHeight - (cmdFun(0).Height + 100) * 4) / 2
    cmdFun(1).Top = cmdFun(0).Top + cmdFun(0).Height + 100
    cmdFun(2).Top = cmdFun(1).Top + cmdFun(1).Height + 100
    cmdFun(3).Top = cmdFun(2).Top + cmdFun(2).Height + 100
End Sub

Private Sub picPageSub_Resize(index As Integer)
    On Error Resume Next
    With tbPageSub(index)
        .Left = 0
        .Top = 0
        .Height = picPageSub(index).ScaleHeight
        .Width = picPageSub(index).ScaleWidth
    End With
End Sub

Private Sub PicUnit_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picUnit(index)
        chkForbidBespeak(index).Left = .ScaleLeft + 30
        chkForbidBespeak(index).Top = .ScaleTop + 30
        
        vsfSelNum(index).Left = .ScaleLeft
        vsfSelNum(index).Width = .ScaleWidth
        vsfSelNum(index).Top = chkForbidBespeak(index).Top + chkForbidBespeak(index).Height
        vsfSelNum(index).Height = .ScaleHeight - vsfSelNum(index).Top
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call SelectedPageChanged(tbPage, Item.index)
End Sub

Private Sub SelectedPageChanged(tbPage As TabControl, ByVal index As Integer)
    '切换页面，对正在编辑的数据进行检查
    '入参:
    '   Index 当前选择Page.Index
    Dim intSelectedGridIndex As Integer
    
    On Error GoTo ErrHandler
    If Val(tbPage.Tag) < tbPage.ItemCount Then
        mblnNotClick = True
        tbPage.Enabled = False
        tbPage.Item(Val(tbPage.Tag)).Selected = True
        tbPage.Enabled = True
        
        intSelectedGridIndex = SelectedVsfGridIndex
        If intSelectedGridIndex >= vsfSelNum.LBound And intSelectedGridIndex <= vsfSelNum.UBound Then
            mblnValiedCanSave = True
            vsfSelNum(intSelectedGridIndex).FinishEditing False    '对正在编辑的数据进行检查
            If mblnValiedCanSave = False Then mblnNotClick = False: Exit Sub
        End If
        
        tbPage.Enabled = False
        tbPage.Item(index).Selected = True
        tbPage.Enabled = True
        mblnNotClick = False
    End If
    
    tbPage.Tag = index
    Call SetButtonEnable(SelectedVsfGridIndex)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnNotClick = False
End Sub

Private Sub tbPageSub_SelectedChanged(index As Integer, ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call SelectedPageChanged(tbPageSub(index), Item.index)
End Sub

Private Sub UserControl_Initialize()
    Call InitFace
    Call SetUnitVisible
End Sub

Private Sub UserControl_Resize()
    Dim sngSplit As Single
    
    Err = 0: On Error Resume Next
    With vsUnit
        .Left = 0
        .Top = optBespeakMode(0).Top + optBespeakMode(0).Height + 50
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - .Top
    End With
    If m_EditMode = ED_RegistPlan_Edit Then
        With vsfNotSelNum
            .Left = vsUnit.Left
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With picFun
            .Left = vsfNotSelNum.Left + vsfNotSelNum.Width + 20
            .Top = vsUnit.Top
            .Height = vsUnit.Height
        End With
        With tbPage
            .Left = IIf(picFun.Tag = "", picFun.Left + picFun.Width, picFun.Left) + 20
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    
        '加一个边框
        sngSplit = 30
        With tbPage
            .Left = .Left + sngSplit
            .Top = .Top + sngSplit
            .Height = .Height - 2 * sngSplit
            .Width = .Width - 2 * sngSplit
        End With
        With shpBack
            .Left = tbPage.Left - sngSplit
            .Top = tbPage.Top - sngSplit
            .Height = tbPage.Height + 2 * sngSplit
            .Width = tbPage.Width + 2 * sngSplit
        End With
    Else
        With tbPage
            .Left = 0
            .Top = vsUnit.Top
            .Height = vsUnit.Height
            .Width = ScaleWidth - .Left
        End With
    End If
End Sub

Private Sub InitUnitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2016-01-11 14:23:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objParentPage As TabControl
    Dim objUnit As 合作单位控制, lngRow As Long
    Dim intPageCount As Integer, p As Integer
    Dim intPage As Integer, intParentPage As Integer
    
    Err = 0: On Error GoTo Errhand:
    intParentPage = tbPage.Selected.index
    intPage = -1
    If Not tbPageSub(intParentPage).Selected Is Nothing Then
        intPage = tbPageSub(intParentPage).Selected.index
    End If
    
    For p = tbPageSub.LBound To tbPageSub.UBound
        tbPageSub(p).RemoveAll
    Next
    intPageCount = picUnit.Count
    
    If Not mobj所有合作单位 Is Nothing Then
        For Each objUnit In mobj所有合作单位
            If lngRow >= intPageCount Then
                Load chkForbidBespeak(lngRow): chkForbidBespeak(lngRow).Visible = True
                Load vsfSelNum(lngRow): vsfSelNum(lngRow).Visible = True
                Load picUnit(lngRow): picUnit(lngRow).Visible = True
                Set chkForbidBespeak(lngRow).Container = picUnit(lngRow)
                Set vsfSelNum(lngRow).Container = picUnit(lngRow)
                picUnit(lngRow).TabStop = False
            End If
            
            picUnit(lngRow).Visible = True
            Set objParentPage = IIf(objUnit.类型 = 1, tbPageSub(Pg_合作单位), tbPageSub(Pg_预约方式))
            Set objItem = objParentPage.InsertItem(objParentPage.ItemCount, objUnit.合作单位名称, picUnit(lngRow).Hwnd, 0)
            objItem.Tag = objUnit.类型 '1-三方机构;2-预约方式
            vsfSelNum(lngRow).Tag = objUnit.类型 '1-三方机构;2-预约方式
            picUnit(lngRow).Tag = objUnit.合作单位名称
            lngRow = lngRow + 1
        Next
    End If
    
    If lngRow <= intPageCount Then
        For i = IIf(lngRow = 0, 1, lngRow) To picUnit.UBound
            Unload chkForbidBespeak(i)
            Unload vsfSelNum(i)
            Unload picUnit(i)
        Next
    End If
    
    '显示/隐藏"合作单位"页签
    If tbPageSub(Pg_合作单位).ItemCount = 0 And intParentPage = Pg_合作单位 Then
        tbPage.Enabled = False: mblnNotClick = True
        tbPage.Item(Pg_预约方式).Selected = True
        tbPage.Enabled = True: mblnNotClick = False
        intParentPage = Pg_预约方式
        
        tbPage(Pg_合作单位).Visible = False
    Else
        tbPage(Pg_合作单位).Visible = True
    End If
    
    '恢复选择页签
    If intPage > tbPageSub(intParentPage).ItemCount - 1 Then
        intPage = tbPageSub(intParentPage).ItemCount - 1
    End If
    If intPage = -1 Then intPage = 0
    
    If tbPageSub(intParentPage).ItemCount > 0 Then
        '手动触发SelectedChanged事件
        Call tbPageSub_SelectedChanged(intParentPage, tbPageSub(intParentPage).Item(intPage))
        tbPageSub(intParentPage).Enabled = False
        tbPageSub(intParentPage).Item(intPage).Selected = True
        tbPageSub(intParentPage).Enabled = True
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnNotClick = False
End Sub

Private Sub InitUnitGrid()
    '初始化合作单位序号网格
    Dim i As Integer
    
    Err = 0: On Error GoTo Errhand:
    With vsfNotSelNum
        .Clear 1
        .Rows = 1
        .HighLight = flexHighlightAlways
        .ColHidden(-1) = False
    End With
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        With vsfSelNum(i)
            .Clear 1
            .Rows = 1
            .Editable = flexEDNone
            .ColHidden(-1) = False
            .HighLight = flexHighlightAlways
            .FocusRect = flexFocusNone
        End With
        mblnNotClick = True
        chkForbidBespeak(i).Value = vbUnchecked
        mblnNotClick = False
    Next
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetUnitVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据合作单位的预约控制方式，设置对应的控件显示
    '编制:刘兴洪
    '日期:2016-01-12 11:23:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    vsfNotSelNum.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value
    picFun.Visible = m_EditMode = ED_RegistPlan_Edit And optBespeakMode(2).Value And Val(picFun.Tag) = 0
    tbPage.Visible = optBespeakMode(2).Value
    vsUnit.Visible = Not optBespeakMode(2).Value
    chkOnlyOneUse.Visible = Not optBespeakMode(2).Value
    optBespeakMode(2).Visible = Val(optBespeakMode(2).Tag) = 0
    If Val(optBespeakMode(2).Tag) = 0 Then
        chkOnlyOneUse.Left = optBespeakMode(2).Left + optBespeakMode(2).Width + 50
    Else
        chkOnlyOneUse.Left = optBespeakMode(2).Left
    End If
    vsUnit.TextMatrix(0, vsUnit.ColIndex("数量")) = IIf(optBespeakMode(0).Value, "比例(%)", "限约数")
    Call UserControl_Resize
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function Get合作单位控制集() As 合作单位控制集
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取合作单信息信息数据
    '返回:号序信息集
    '编制:刘兴洪
    '日期:2016-01-13 12:34:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, intCol As Integer, index As Integer
    Dim objUnits As New 合作单位控制集, objUnit As 合作单位控制
    Dim lngSum As Double, varTemp As Variant
    Dim strUnitName As String
    Dim objNums As 号序信息集, objNum As 号序信息
    Dim bln禁止预约 As Boolean
    
    Err = 0: On Error GoTo Errhand:
    '数据未改变，直接返回原集合的副本
    If m_IsDataChanged = False Then
        If mobj合作单位集.Count = 0 And mobj所有合作单位.Count > 0 Then
            '第一次加载，没有改变，应该是全部不限制
            
        Else
            Set Get合作单位控制集 = mobj合作单位集.Clone
            Exit Function
        End If
    End If
    
    '数据已改变，重新构造集合对象
    With objUnits
        .预约控制方式 = GetSelectedIndex(optBespeakMode) + 1
        .是否独占 = chkOnlyOneUse.Value = vbChecked
        .是否修改 = True
    End With
    
    If optBespeakMode(0).Value Or optBespeakMode(1).Value Then
        '按比例控制或按总量控制
        With vsUnit
            For lngRow = 1 To .Rows - 1
                If .IsSubtotal(lngRow) = False Then
                    Set objUnit = New 合作单位控制
                    objUnit.合作单位名称 = .TextMatrix(lngRow, .ColIndex("合作单位"))
                    objUnit.类型 = .RowData(lngRow)
                    
                    If .RowHidden(lngRow) Then '隐藏的就是禁止预约
                        bln禁止预约 = True
                        lngSum = 0
                    Else
                        bln禁止预约 = Abs(Val(.TextMatrix(lngRow, .ColIndex("禁止预约")))) = 1
                        lngSum = Val(.TextMatrix(lngRow, .ColIndex("数量")))
                    End If
                    '0-禁止预约;1-按比例控制预约;2-按总量控制预约;3-按序号控制预约;4-不作限制
                    objUnit.预约控制方式 = IIf(bln禁止预约, 0, _
                                            IIf(lngSum = 0, 4, IIf(optBespeakMode(0).Value, 1, 2)))
                    Set objNums = New 号序信息集
                    If lngSum > 0 Or bln禁止预约 Then
                        Set objNum = New 号序信息
                        objNum.序号 = 0
                        objNum.数量 = lngSum
                        objNums.AddItem objNum
                    End If
                    Set objUnit.号序信息集 = objNums
                    objUnits.AddItem objUnit, "K" & objUnit.类型 & "_" & objUnit.合作单位名称
                End If
            Next
        End With
    Else
        '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
        For index = vsfSelNum.LBound To vsfSelNum.UBound
            If GetLocaleUnit(index, objUnit) Then
                objUnits.AddItem objUnit, "K" & objUnit.类型 & "_" & objUnit.合作单位名称
            End If
        Next
    End If
    Set Get合作单位控制集 = objUnits
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetLocaleUnit(ByVal index As Integer, ByRef objUnit As 合作单位控制) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的合作单信息
    '入参:index-指定的页
    '出参:objUnit-合作单位信息集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-01-13 18:32:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim objNums As 号序信息集, objNum As 号序信息
    Dim varTemp As Variant, lngCount As Long
    
    Set objUnit = New 合作单位控制
    Err = 0: On Error GoTo Errhand:
    objUnit.合作单位名称 = picUnit(index).Tag
    objUnit.类型 = Val(vsfSelNum(index).Tag)
    '0-禁止预约;1-按比例控制预约;2-按总量控制预约;3-按序号控制预约;4-不作限制
    If chkForbidBespeak(index).Value = vbChecked _
        Or mobj所有号序集.预约控制 = 2 And objUnit.类型 = 1 Then
        '仅禁止三方合作单位
        objUnit.预约控制方式 = 0
    Else
        objUnit.预约控制方式 = 3
    End If

    Set objNums = New 号序信息集
    If objUnit.预约控制方式 = 3 Then
        With vsfSelNum(index)
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("数量"))) <> 0 Then
                    lngCount = lngCount + Val(.TextMatrix(i, .ColIndex("数量")))
                    
                    Set objNum = New 号序信息
                    objNum.序号 = Val(.TextMatrix(i, .ColIndex("序号")))
                    If .TextMatrix(i, .ColIndex("时间段")) <> "" Then
                        varTemp = Split(.Cell(flexcpData, i, .ColIndex("时间段")), "-")
                        objNum.开始时间 = varTemp(0)
                        objNum.终止时间 = varTemp(1)
                    End If
                    objNum.数量 = Val(.TextMatrix(i, .ColIndex("数量")))
                    objNums.AddItem objNum
                End If
            Next
        End With
        '一个序号都没有设置数量,则表示不限制
        If lngCount = 0 Then objUnit.预约控制方式 = 4
    End If
    If objUnit.预约控制方式 = 0 Or objUnit.预约控制方式 = 4 Then
        '禁止预约或不限制时添加一个记录，以便保存
        Set objNum = New 号序信息
        objNum.序号 = 0
        objNum.数量 = 0
        objNums.AddItem objNum
    End If
    Set objUnit.号序信息集 = objNums
    GetLocaleUnit = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Property Get Get合作单位控制信息集() As 合作单位控制集
   Set Get合作单位控制信息集 = Get合作单位控制集
End Property

Public Function IsValied(Optional ByVal blnChanged As Boolean) As Boolean
    '检查数据
    '外面一层是否改变，若改变则本层也要检查
    Dim lngSum As Double, lng限约数 As Long, lngSN As Long
    Dim i As Long, j As Integer, k As Long
    Dim intSelectedGridIndex As Integer
    
    Err = 0: On Error GoTo ErrHandler
    '数据未改变不检查
    If m_IsDataChanged = False And blnChanged = False Then IsValied = True: Exit Function
    
    mblnValiedCanSave = True
    vsUnit.FinishEditing False
    If mblnValiedCanSave = False Then Exit Function
    
    mblnValiedCanSave = True
    intSelectedGridIndex = SelectedVsfGridIndex
    If intSelectedGridIndex >= vsfSelNum.LBound And intSelectedGridIndex <= vsfSelNum.UBound Then
        vsfSelNum(intSelectedGridIndex).FinishEditing False
    End If
    If mblnValiedCanSave = False Then Exit Function

    If optBespeakMode(0).Value Then '按比例
        If chkOnlyOneUse.Value = vbChecked Then
            For i = 1 To vsUnit.Rows - 1
                If vsUnit.IsSubtotal(i) = False Then
                    lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")))
                End If
            Next
            If lngSum > 100 Then
                MsgBox "独占方式时，限约比例之和不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        Else
            For i = 1 To vsUnit.Rows - 1
                If vsUnit.IsSubtotal(i) = False Then
                    lngSum = Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")))
                    If lngSum > 100 Then
                        MsgBox vsUnit.TextMatrix(i, vsUnit.ColIndex("合作单位")) & "的预约比例不能超过100！", vbInformation + vbOKOnly, gstrSysName
                        vsUnit.Row = i: vsUnit.Col = vsUnit.ColIndex("数量")
                        Exit Function
                    End If
                End If
            Next
        End If
    ElseIf optBespeakMode(1).Value Then '按总量
        If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
        If lng限约数 > 0 Then '不限约时不用检查
            If chkOnlyOneUse.Value = vbChecked Then
                For i = 1 To vsUnit.Rows - 1
                    If vsUnit.IsSubtotal(i) = False Then
                        lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")))
                    End If
                Next
                If lngSum > lng限约数 Then
                    MsgBox "独占方式时，限约数之和不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            Else
                For i = 1 To vsUnit.Rows - 1
                    If vsUnit.IsSubtotal(i) = False Then
                        lngSum = Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")))
                        If lngSum > lng限约数 Then
                            MsgBox vsUnit.TextMatrix(i, vsUnit.ColIndex("合作单位")) & "的限约数不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                            vsUnit.Row = i: vsUnit.Col = vsUnit.ColIndex("数量")
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    Else '按序号
        If Not mobj所有号序集 Is Nothing Then
            If mobj所有号序集.是否分时段 And mobj所有号序集.是否序号控制 = False Then
                For k = 1 To vsfNotSelNum.Rows - 1
                    lngSum = Val(vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("数量")))
                    lngSN = Val(vsfNotSelNum.TextMatrix(k, vsfNotSelNum.ColIndex("序号")))
                    For i = vsfSelNum.LBound To vsfSelNum.UBound
                        For j = 1 To vsfSelNum(i).Rows - 1
                            If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号"))) = lngSN Then
                                lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("数量")))
                            End If
                        Next
                    Next
                    If lngSum < 0 Then
                        MsgBox vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("时间段")) & _
                            " 分配的预约数超过了该时段的可预约数量(" & Val(vsfNotSelNum.Cell(flexcpData, k, vsfNotSelNum.ColIndex("数量"))) & ")！", _
                            vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                Next
            End If
        End If
    End If
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "返回/设置对象中文本和图形的背景色。"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    On Error Resume Next
    lblEdit.BackColor = UserControl.BackColor
    optBespeakMode(0).BackColor = UserControl.BackColor
    optBespeakMode(1).BackColor = UserControl.BackColor
    optBespeakMode(2).BackColor = UserControl.BackColor
    chkOnlyOneUse.BackColor = UserControl.BackColor
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "指出 Label 或 Shape 的背景样式是透明的还是不透明的。"
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_IsDataChanged = m_def_IsDataChanged
    m_EditMode = m_def_EditMode
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    m_IsDataChanged = PropBag.ReadProperty("IsDataChanged", m_def_IsDataChanged)
    m_EditMode = PropBag.ReadProperty("EditMode", m_def_EditMode)
End Sub

Private Sub UserControl_Terminate()
    Set mobj合作单位集 = Nothing
    Set mobj所有号序集 = Nothing
    Set mobj所有合作单位 = Nothing
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IsDataChanged", m_IsDataChanged, m_def_IsDataChanged)
    Call PropBag.WriteProperty("EditMode", m_EditMode, m_def_EditMode)
End Sub

Private Sub vsfNotSelNum_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetButtonEnable(SelectedVsfGridIndex)
End Sub

Private Sub vsfNotSelNum_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
End Sub

Private Sub vsfNotSelNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfNotSelNum_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
End Sub

Private Sub vsfSelNum_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    RemoveItem vsfNotSelNum, vsfSelNum(index), _
        Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("序号"))), _
        True, Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("数量")))
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsfSelNum_AfterRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    Call SetButtonEnable(index)
    If vsfSelNum(index).Editable = flexEDKbdMouse Then
        vsfNotSelNum.Row = NewRow
        vsfNotSelNum.TopRow = vsfSelNum(index).TopRow
    End If
End Sub

Private Sub vsfSelNum_BeforeEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If vsfSelNum(index).ColIndex("数量") <> Col Then Cancel = True: Exit Sub
End Sub

Private Sub vsfSelNum_EnterCell(index As Integer)
    If vsfSelNum(index).Col = vsfSelNum(index).ColIndex("数量") Then vsfSelNum(index).EditCell
End Sub

Private Sub vsfSelNum_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And vsfSelNum(index).Editable = flexEDKbdMouse Then
        If vsfSelNum(index).Row = vsfSelNum(index).Rows - 1 And vsfSelNum(index).Col = vsfSelNum(index).Cols - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsfSelNum(index), 2)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsfSelNum_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsfSelNum_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '输入位数限制，整数位长度不能大于9
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsfSelNum(index).EditText) >= 9 Then KeyAscii = 0
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsfSelNum_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Long, lngSN As Long
    Dim i As Integer, j As Integer, lngRow As Long
    
    On Error GoTo Errhand
    '整数位多余9位的直接截掉,防止溢出
    If InStr(vsfSelNum(index).EditText, ".") > 0 Then
        If InStr(vsfSelNum(index).EditText, ".") > 9 Then
            vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
        End If
    Else
        vsfSelNum(index).EditText = Left(vsfSelNum(index).EditText, 9)
    End If
    
    lngSN = Val(vsfSelNum(index).TextMatrix(Row, vsfSelNum(index).ColIndex("序号")))
    For i = 1 To vsfNotSelNum.Rows - 1
        If Val(vsfNotSelNum.TextMatrix(i, vsfNotSelNum.ColIndex("序号"))) = lngSN Then
            lngRow = i: Exit For
        End If
    Next
    '计算剩余数量
    lngSum = Val(vsfNotSelNum.Cell(flexcpData, lngRow, vsfNotSelNum.ColIndex("数量")))
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        If i <> index Then
            For j = 1 To vsfSelNum(i).Rows - 1
                If Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号"))) = lngSN Then
                    lngSum = lngSum - Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("数量")))
                    Exit For
                End If
            Next
        End If
    Next
    
    If Val(vsfSelNum(index).EditText) > lngSum Then
        MsgBox picUnit(index).Tag & " 预约数(" & Val(vsfSelNum(index).EditText) & ")不能超过剩余预约数量(" & lngSum & ")！", vbInformation + vbOKOnly, gstrSysName
        Cancel = True: mblnValiedCanSave = False: Exit Sub
    End If
    vsfSelNum(index).EditText = FormatEx(Val(vsfSelNum(index).EditText), 0)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsUnit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = vsUnit.ColIndex("禁止预约") Then
        If vsUnit.TextMatrix(Row, vsUnit.ColIndex("禁止预约")) = True Then
            vsUnit.TextMatrix(Row, vsUnit.ColIndex("数量")) = ""
            vsUnit.Cell(flexcpBackColor, Row, vsUnit.ColIndex("数量")) = vbButtonFace
        Else
            vsUnit.Cell(flexcpBackColor, Row, vsUnit.ColIndex("数量")) = vsUnit.BackColor
        End If
    End If
End Sub

Private Sub vsUnit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If m_EditMode <> ED_RegistPlan_Edit Then Cancel = True: Exit Sub
    If vsUnit.IsSubtotal(Row) Then Cancel = True: Exit Sub
    If Col = vsUnit.ColIndex("合作单位") Then Cancel = True: Exit Sub
    If Col = vsUnit.ColIndex("数量") Then
        If Abs(Val(vsUnit.TextMatrix(Row, vsUnit.ColIndex("禁止预约")))) = 1 Then Cancel = True: Exit Sub
    End If
    '由事件AfterEdit调到这里，因为当正在编辑时直接按保存，检查不到
    m_IsDataChanged = True: RaiseEvent DataIsChanged
End Sub

Private Sub vsUnit_EnterCell()
    If vsUnit.Col = vsUnit.ColIndex("数量") Then vsUnit.EditCell
End Sub

Private Sub vsUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If vsUnit.Row = vsUnit.Rows - 1 And vsUnit.Col = vsUnit.Cols - 1 Then
            'Call zlCommFun.PressKey(vbKeyTab)
        Else
            Call zlVsMoveGridCell(vsUnit, 1)
        End If
        KeyCode = 0
    End If
End Sub

Private Sub vsUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub vsUnit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack Then Exit Sub
    '输入位数限制，整数位长度不能大于9
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then KeyAscii = 0
    Else
        If Len(vsUnit.EditText) >= 9 Then KeyAscii = 0
    End If
    If optBespeakMode(0).Value Then
        If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Else
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Public Property Let 预约控制(ByVal vNewValue As Byte)
    Dim i As Integer, j As Integer
    
    On Error GoTo Errhand
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    mobj所有号序集.预约控制 = vNewValue
    '0-不作预约限制;1-该号别禁止预约;2-仅禁止三方机构平台的预约
    Call UnitPageVisible(mobj所有号序集.预约控制 <> 2)
    
    '清除数据
    If mobj所有号序集.预约控制 = 2 Then
        For i = 1 To vsUnit.Rows - 1
            If vsUnit.RowData(i) = 1 And vsUnit.IsSubtotal(i) = False Then
                vsUnit.TextMatrix(i, vsUnit.ColIndex("禁止预约")) = 1
                vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")) = ""
                vsUnit.Cell(flexcpBackColor, i, vsUnit.ColIndex("数量")) = vbButtonFace
            End If
        Next
        For i = vsfSelNum.LBound To vsfSelNum.UBound
            If Val(vsfSelNum(i).Tag) = Val("1-三方结构") Then
                chkForbidBespeak(i).Value = vbChecked
                For j = 1 To vsfSelNum(i).Rows - 1
                    If mobj所有号序集.是否序号控制 Then
                        If j > vsfSelNum(i).Rows - 1 Then Exit For
                        RemoveItem vsfSelNum(i), vsfNotSelNum, Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号")))
                        j = j - 1
                    Else
                        RemoveItem vsfNotSelNum, vsfSelNum(i), Val(vsfSelNum(i).TextMatrix(j, vsfSelNum(i).ColIndex("序号"))), True, 0
                    End If
                Next
            End If
        Next
        Call SetButtonEnable(SelectedVsfGridIndex)
    End If
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Public Property Let 所有号序信息集(ByVal vNewValue As 号序信息集)
    Err = 0: On Error GoTo Errhand
    Set mobj所有号序集 = vNewValue
    If mobj所有号序集 Is Nothing Then Set mobj所有号序集 = New 号序信息集
    Set mobj合作单位集 = Get合作单位控制集
    Call InitData
    Exit Property
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Property

Private Sub vsUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngSum As Double, lng限约数 As Long
    Dim i As Long
    
    On Error GoTo Errhand
    '编辑禁止预约列时不检查
    If Col = vsUnit.ColIndex("禁止预约") Then Exit Sub
    '整数位多余9位的直接截掉,防止溢出
    If InStr(vsUnit.EditText, ".") > 0 Then
        If InStr(vsUnit.EditText, ".") > 9 Then
            vsUnit.EditText = Left(vsUnit.EditText, 9)
        End If
    Else
        vsUnit.EditText = Left(vsUnit.EditText, 9)
    End If
    
    If chkOnlyOneUse.Value = vbChecked Then
        For i = 1 To vsUnit.Rows - 1
            If i <> vsUnit.Row And vsUnit.IsSubtotal(i) = False Then
                lngSum = lngSum + Val(vsUnit.TextMatrix(i, vsUnit.ColIndex("数量")))
            End If
        Next
        lngSum = lngSum + Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '按比例
            If lngSum > 100 Then
                MsgBox "独占方式时，合作单位控制限约比例之和不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then '按总量
            If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
            If lng限约数 > 0 Then '不限约时不用检查
                If lngSum > lng限约数 Then
                    MsgBox "独占方式时，合作单位控制限约数之和不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    Else
        lngSum = Val(vsUnit.EditText)
        If optBespeakMode(0).Value Then '按比例
            If lngSum > 100 Then
                MsgBox vsUnit.TextMatrix(vsUnit.Row, vsUnit.ColIndex("合作单位")) & " 预约比例不能超过100！", vbInformation + vbOKOnly, gstrSysName
                Cancel = True: mblnValiedCanSave = False: Exit Sub
            End If
        ElseIf optBespeakMode(1).Value Then  '按总量
            If Not mobj所有号序集 Is Nothing Then lng限约数 = mobj所有号序集.限约数
            If lng限约数 > 0 Then '不限约时不用检查
                If lngSum > lng限约数 Then
                    MsgBox vsUnit.TextMatrix(vsUnit.Row, vsUnit.ColIndex("合作单位")) & " 限约数不能超过限约数(" & lng限约数 & ")！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True: mblnValiedCanSave = False: Exit Sub
                End If
            End If
        End If
    End If
    vsUnit.EditText = FormatEx(Val(vsUnit.EditText), 2)
    vsUnit.EditText = IIf(vsUnit.EditText = "0", "", vsUnit.EditText)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,false
Public Property Get IsDataChanged() As Boolean
    IsDataChanged = m_IsDataChanged
End Property

Public Property Let IsDataChanged(ByVal New_IsDataChanged As Boolean)
    m_IsDataChanged = New_IsDataChanged
    PropertyChanged "IsDataChanged"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get EditMode() As gRegistPlanEditMode
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal New_EditMode As gRegistPlanEditMode)
    Dim i As Integer
    
    m_EditMode = IIf(New_EditMode = ED_RegistPlan_UpdateUnit, ED_RegistPlan_Edit, New_EditMode)
    If mobj所有号序集 Is Nothing Then
        m_EditMode = ED_RegistPlan_View
    ElseIf m_EditMode = ED_RegistPlan_Edit And mobj所有号序集.预约控制 = Val("1-禁止预约") Then
        m_EditMode = ED_RegistPlan_View
    End If
    PropertyChanged "EditMode"
    
    For i = optBespeakMode.LBound To optBespeakMode.UBound
        optBespeakMode(i).Enabled = m_EditMode = ED_RegistPlan_Edit
    Next
    chkOnlyOneUse.Enabled = m_EditMode = ED_RegistPlan_Edit
    vsUnit.Editable = IIf(m_EditMode = ED_RegistPlan_Edit, flexEDKbdMouse, flexEDNone)
    picFun.Enabled = m_EditMode = ED_RegistPlan_Edit
    For i = vsfSelNum.LBound To vsfSelNum.UBound
        chkForbidBespeak(i).Enabled = m_EditMode = ED_RegistPlan_Edit
        vsfSelNum(i).Editable = IIf(m_EditMode = ED_RegistPlan_Edit And chkForbidBespeak(i).Value = Unchecked And mblnEdit, flexEDKbdMouse, flexEDNone)
    Next
    
    '隐藏
    SetUnitVisible
    UserControl_Resize
End Property

