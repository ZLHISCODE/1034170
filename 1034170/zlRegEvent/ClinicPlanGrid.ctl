VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl ClinicPlanGrid 
   BackColor       =   &H80000005&
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11445
   ScaleHeight     =   6885
   ScaleWidth      =   11445
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   4305
      Left            =   630
      TabIndex        =   0
      Top             =   240
      Width           =   6555
      _cx             =   11562
      _cy             =   7594
      Appearance      =   0
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanGrid.ctx":0000
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
      Begin VB.PictureBox Picture1 
         Height          =   105
         Left            =   5400
         ScaleHeight     =   105
         ScaleWidth      =   30
         TabIndex        =   5
         Top             =   2160
         Width           =   30
      End
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   30
         Picture         =   "ClinicPlanGrid.ctx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   4
         Top             =   90
         Width           =   150
      End
      Begin VB.PictureBox picToolTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   630
         ScaleHeight     =   540
         ScaleWidth      =   1365
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1110
         Visible         =   0   'False
         Width           =   1395
         Begin VB.Label lblToolTip1 
            BackStyle       =   0  'Transparent
            Caption         =   "已挂数：120"
            Height          =   180
            Left            =   90
            TabIndex        =   3
            Top             =   60
            Width           =   1290
         End
         Begin VB.Label lblToolTip2 
            BackStyle       =   0  'Transparent
            Caption         =   "限号数：220"
            Height          =   180
            Left            =   90
            TabIndex        =   2
            Top             =   300
            Width           =   1290
         End
      End
   End
   Begin VB.Image imgReplace 
      Height          =   120
      Left            =   9270
      Picture         =   "ClinicPlanGrid.ctx":016B
      Top             =   2340
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   6435
      Left            =   480
      Top             =   90
      Width           =   10425
   End
   Begin VB.Image imgLock 
      Height          =   210
      Left            =   8880
      Picture         =   "ClinicPlanGrid.ctx":0615
      Top             =   2310
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "ClinicPlanGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Enum mPlanGridFixedColIndex '网格固定列索引
    COl_空列 = 0
    COL_号源ID
    COL_安排ID
    COL_号类
    col_号码
    COL_项目
    COL_科室
    Col_医生
    
    COL_开始时间
    COL_终止时间
    
    COL_是否建病案
    COL_是否序号控制
    COL_是否分时段
    COL_预约天数
    COL_出诊频次
    COL_假日控制状态
    COL_排班方式
    COL_分诊方式
End Enum

Public Enum gDataStyle
    Data_Templet = 0
    Data_FixedRule = 1
    Data_Plan = 2
End Enum

Private Const m_def_DataStyle = 0

Private m_DataStyle As gDataStyle
Private m_MinDate As Date
Private m_MaxDate As Date

'事件声明
Public Event AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mrsData As ADODB.Recordset

Public Property Get vsfGrid() As VSFlexGrid
    Set vsfGrid = vsfRegistPlan
End Property

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

Public Property Get DataStyle() As gDataStyle
Attribute DataStyle.VB_Description = "返回/设置表格类型。"
    DataStyle = m_DataStyle
End Property

Public Property Let DataStyle(ByVal New_DataStyle As gDataStyle)
    m_DataStyle = New_DataStyle
    PropertyChanged "DataStyle"
    Call InitPlanGrid
End Property

Private Sub UserControl_Initialize()
    Call InitPlanGrid
    Call SetPicToolTipEffect
End Sub

Private Sub UserControl_InitProperties()
    m_DataStyle = m_def_DataStyle
    m_MinDate = Format(Now, "yyyy-mm-dd")
    m_MaxDate = Format(Now, "yyyy-mm-dd")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DataStyle = PropBag.ReadProperty("DataStyle", m_def_DataStyle)
    m_MinDate = PropBag.ReadProperty("MinDate", Format(Now, "yyyy-mm-dd"))
    m_MaxDate = PropBag.ReadProperty("MaxDate", Format(Now, "yyyy-mm-dd"))
    Call InitPlanGrid
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    vsfRegistPlan.Move 10, 10, UserControl.ScaleWidth - 20, UserControl.ScaleHeight - 20
End Sub

Private Sub UserControl_Terminate()
    Set mrsData = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DataStyle", m_DataStyle, m_def_DataStyle)
    Call PropBag.WriteProperty("MinDate", m_MinDate, Format(Now, "yyyy-mm-dd"))
    Call PropBag.WriteProperty("MaxDate", m_MaxDate, Format(Now, "yyyy-mm-dd"))
End Sub

'开始时间
Public Property Get MinDate() As Date
    MinDate = m_MinDate
End Property

Public Property Let MinDate(ByVal vNewValue As Date)
    m_MinDate = Format(vNewValue, "yyyy-mm-dd")
    If m_MinDate > m_MaxDate Then m_MaxDate = m_MinDate
    PropertyChanged "MinDate"
    Call InitPlanGrid
End Property

'结束时间
Public Property Get MaxDate() As Date
    MaxDate = m_MaxDate
End Property

Public Property Let MaxDate(ByVal vNewValue As Date)
    Dim dtCur As Date
    dtCur = Format(vNewValue, "yyyy-mm-dd")
    If m_MinDate > dtCur Then Exit Property
    m_MaxDate = dtCur
    PropertyChanged "MaxDate"
    Call InitPlanGrid
End Property

'当前选择项属性
'------------------------------------------
Public Property Get 限制项目() As String
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        限制项目 = .Cell(flexcpData, 0, lngCol)
    End With
End Property

Public Property Get 时间段() As String
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        时间段 = .TextMatrix(.Row, lngCol)
    End With
End Property

Public Property Get 号源ID() As Long
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        If .RowData(.Row) = -1 Then Exit Property '空行
        号源ID = Val(.TextMatrix(.Row, COL_号源ID))
    End With
End Property

Public Property Get 安排ID() As Long
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        If .RowData(.Row) = -1 Then Exit Property '空行
        安排ID = Val(.TextMatrix(.Row, COL_安排ID))
    End With
End Property

Public Property Get 记录ID() As Long
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        记录ID = Val(.Cell(flexcpData, .Row, lngCol))
    End With
End Property

Public Property Get Is锁号() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        Is锁号 = Not .Cell(flexcpPicture, .Row, lngCol) Is Nothing
    End With
End Property

Public Property Get Is停诊() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        Is停诊 = .Cell(flexcpForeColor, .Row, lngCol) = vbRed
    End With
End Property

Public Property Get Is替诊() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        Is替诊 = .Cell(flexcpForeColor, .Row, lngCol) = vbBlue
    End With
End Property

Public Property Get IsSelectedNotNull() As Boolean
    Dim lngCol As Long
    
    '判断当前选择是否为空
    Err = 0: On Error GoTo errHandler
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        lngCol = GetItemNameCol(.Col, .FixedCols) '时段列
        If .RowData(.Row) = -1 Or Trim(.TextMatrix(.Row, lngCol)) = "" Then Exit Property '空行或无时段列
    End With
    IsSelectedNotNull = True
    Exit Property
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property
'------------------------------------------

Private Sub SetPicToolTipEffect()
    '功能：设置提示框PicToolTip的边框
    Dim lngR As Long
    
    '边框：API=RoundRect
    picToolTip.Line (Screen.TwipsPerPixelX, 0)-(picToolTip.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    picToolTip.Line (Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY)-(picToolTip.Width - Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.Line (0, Screen.TwipsPerPixelY)-(0, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.Line (picToolTip.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(picToolTip.Width - Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    picToolTip.PSet (picToolTip.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    picToolTip.PSet (Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    picToolTip.PSet (picToolTip.Width - Screen.TwipsPerPixelX * 2, picToolTip.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    
    '形状
    lngR = CreateRoundRectRgn(0, 0, picToolTip.ScaleX(picToolTip.Width, picToolTip.ScaleMode, vbPixels) + 1, picToolTip.ScaleY(picToolTip.Height, picToolTip.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(picToolTip.Hwnd, lngR, False)
End Sub

Private Sub InitPlanGrid()
    '功能：初始化安排数据表格
    '   vsfGrid - VSF表格
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant
    Dim dtCurDate As Date, dtMaxDate As Date, intDays As Integer

    Err = 0: On Error GoTo errHandler
    With vsfRegistPlan
        .Redraw = False
        .Rows = 2
        
        '固定列
        strHead = " ,4,200|号源ID,4,0|安排ID,4,0|号类,4,500|号码,4,500|项目,1,1000|科室,1,1000|医生,1,650"
        strHeadSub = " ,号源ID,安排ID,号类,号码,项目,科室,医生"
        If m_DataStyle = Data_FixedRule Then
            strHead = strHead & "|开始时间,4,1900|终止时间,4,1900"
            strHeadSub = strHeadSub & ",开始时间,终止时间"
        End If
        varData = Split(strHead, "|")
        varDataSub = Split(strHeadSub, ",")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0): .TextMatrix(1, i) = varDataSub(i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = .Cols: .FixedRows = 2
        '动态列
        Select Case m_DataStyle
        Case Data_Templet, Data_FixedRule '模板,固定规则
            strHead = "周一,1,450|周一,4,550|周一,4,550|周二,1,450|周二,4,450|周二,4,550|周三,1,450|周三,4,550|周三,4,550|" & _
                    "周四,1,450|周四,4,550|周四,4,550|周五,1,450|周五,4,550|周五,4,550|周六,1,450|周六,4,550|周六,4,550|" & _
                    "周日,1,450|周日,4,550|周日,4,550"
            strHeadSub = "时段,限号,限约,时段,限号,限约,时段,限号,限约," & _
                    "时段,限号,限约,时段,限号,限约,时段,限号,限约," & _
                    "时段,限号,限约"
            If m_DataStyle = Data_Templet Then
                strHead = strHead & "|其他规则,1,1150|其他规则,1,450|其他规则,4,550|其他规则,4,550"
                strHeadSub = strHeadSub & ",限制项目,时段,限号,限约"
            End If
            varData = Split(strHead, "|")
            varDataSub = Split(strHeadSub, ",")
            lngCol = .Cols
            .Cols = .Cols + UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, lngCol) = Split(varData(i), ",")(0): .TextMatrix(1, lngCol) = varDataSub(i)
                .Cell(flexcpData, 0, lngCol) = CStr(Split(varData(i), ",")(0))
                .ColAlignment(lngCol) = Split(varData(i), ",")(1)
                .ColWidth(lngCol) = Split(varData(i), ",")(2)
'                .ColKey(i) = Split(varData(i), ",")(0) & "-" & varDataSub(i)
                If i Mod 3 <> 0 Then .FixedAlignment(i) = flexAlignCenterCenter
                lngCol = lngCol + 1
            Next
        Case Data_Plan '安排记录
            intDays = DateDiff("d", m_MinDate, m_MaxDate) + 1 '天数
            dtCurDate = m_MinDate
            lngCol = .Cols
            .Cols = .Cols + intDays * 3
            For i = 1 To intDays
                .Cell(flexcpText, 0, lngCol, 0, lngCol + 2) = Format(dtCurDate, "mm月dd日") & Chr(10) & _
                    Choose(Weekday(dtCurDate, vbMonday), "周一", "周二", "周三", "周四", "周五", "周六", "周日")
                .Cell(flexcpData, 0, lngCol) = dtCurDate '日期
                .Cell(flexcpText, 1, lngCol, 1, lngCol + 2) = "时段" & vbTab & "限号" & vbTab & "限约"
                .ColAlignment(lngCol) = 1: .ColAlignment(lngCol + 1) = 4: .ColAlignment(lngCol + 2) = 4
                .ColWidth(lngCol) = 580: .ColWidth(lngCol + 1) = 650: .ColWidth(lngCol + 2) = 650
'                .ColKey(lngCol) = i & "-时段": .ColKey(lngCol + 1) = i & "-限号": .ColKey(lngCol + 2) = i & "-限约"
                .FixedAlignment(lngCol + 1) = flexAlignCenterCenter: .FixedAlignment(lngCol + 2) = flexAlignCenterCenter
                dtCurDate = DateAdd("d", 1, dtCurDate)
                lngCol = lngCol + 3
            Next
            .RowHeight(0) = 500
        End Select
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
'        .WordWrap = True '允许自动换行
        .RowHeightMin = 450
        
        '列属性设置,用于用户选择显示列
        For i = 0 To .Cols - 1
            'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)|列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
            Select Case i
            Case COl_空列, COL_号源ID, COL_安排ID
                 .ColData(i) = "-1|1"
            Case col_号码
                .ColData(i) = "1|0"
            End Select
        Next

        '合并设置
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True: .MergeCol(-1) = True
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadDataByRecordset(ByVal rsData As ADODB.Recordset) As Boolean
    '功能：根据Recordset对象加载数据
    '说明：数据必须是按"号类,号码,科室,项目,医生"进行排序了的，否则可能显示不正确
    Dim i As Long, j As Long, lngCurRow As Long, lngCurCol As Long
    Dim strGroupKey As String '用于纵向按"号类,号码,科室,项目,医生"分组
    Dim lngBackColor As Long '设置纵向组的交替色
    Dim strTemp As String, blnAddRow As Boolean
    Dim lngRowStart As Long, lngRowEnd As Long
    Dim lngOldRow As Long, lngoldCol As Long, blnFind As Boolean
    
    Err = 0: On Error GoTo errHandle
    vsfGrid.Rows = 2
    Set mrsData = Nothing
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    Set mrsData = rsData
    
    With vsfRegistPlan
        lngOldRow = .Row
        lngoldCol = .Col
        lngCurRow = 2
        strGroupKey = ""
        lngBackColor = G_AlternateColor
        Do While Not rsData.EOF
            blnFind = False
            '1.纵向分组
            strTemp = Nvl(rsData!号类) & "," & Nvl(rsData!号码) & "," & Nvl(rsData!科室) & "," & _
                    Nvl(rsData!收费项目) & "," & Nvl(rsData!医生姓名)
            If m_DataStyle = Data_FixedRule Then strTemp = strTemp & "," & Nvl(rsData!开始时间) & "," & Nvl(rsData!终止时间)
            If strGroupKey <> strTemp Then
                strGroupKey = strTemp
                lngCurCol = -1  '用以判断是否确定了列
                lngBackColor = IIf(lngBackColor = vbWindowBackground, G_AlternateColor, vbWindowBackground)
                
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowHidden(lngCurRow) = True '加一行空行，并隐藏
                .RowData(lngCurRow) = -1 '标记，用于判断是否为隐藏空行
                
                lngCurRow = lngCurRow + 1
            End If
            '2.横向分组
            '2.1确定当前列
            Select Case m_DataStyle
            Case Data_Templet '模板
                If Nvl(rsData!排班规则) <> 1 Then
                    '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                    lngCurCol = .Cols - 4: blnFind = True
                Else
                    strTemp = Nvl(rsData!限制项目, "※")
                    For i = .FixedCols To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFind = True
                            Exit For
                        End If
                    Next
                End If
            Case Data_FixedRule '固定规则
                strTemp = Nvl(rsData!限制项目, "※")
                For i = .FixedCols To .Cols - 1 Step 3
                    If strTemp = .Cell(flexcpData, 0, i) Then
                        lngCurCol = i: blnFind = True
                        Exit For
                    End If
                Next
            Case Else '安排记录
                strTemp = Format(Nvl(rsData!出诊日期), "yyyy-mm-dd")
                For i = .FixedCols To .Cols - 1 Step 3
                    If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                        lngCurCol = i: blnFind = True
                        Exit For
                    End If
                Next
            End Select
            If blnFind Then
                '2.2确定当前行
                If m_DataStyle = Data_Templet Then
                    '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                    If Nvl(rsData!排班规则) = 6 Then
                        Call GetGroupRange(IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1), lngRowStart, lngRowEnd)
                        lngCurRow = lngRowEnd + 1
                        For i = lngRowStart To lngRowEnd
                            If .TextMatrix(i, lngCurCol + 1) = "" Then lngCurRow = i: Exit For
                            If .TextMatrix(i, lngCurCol + 1) = Nvl(rsData!上班时段) Then
                                lngCurRow = i: Exit For
                            End If
                        Next
                    Else
                        For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                            If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '是隐藏空行或者无数据行
                                lngCurRow = i + 1: Exit For
                            End If
                        Next
                    End If
                Else
                    For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                        If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '是隐藏空行或者无数据行
                            lngCurRow = i + 1: Exit For
                        End If
                    Next
                End If
            End If
            '3.加载数据
            blnAddRow = False
            If .Rows - 1 < lngCurRow Then
                '已有行不够，加1行
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = lngBackColor '用于设置交替色
                
                .TextMatrix(lngCurRow, COL_号源ID) = Nvl(rsData!号源ID)
                .TextMatrix(lngCurRow, COL_安排ID) = Nvl(rsData!安排ID)
                .TextMatrix(lngCurRow, COL_号类) = Nvl(rsData!号类)
                .TextMatrix(lngCurRow, col_号码) = Nvl(rsData!号码)
                .TextMatrix(lngCurRow, COL_科室) = Nvl(rsData!科室)
                .TextMatrix(lngCurRow, COL_项目) = Nvl(rsData!收费项目)
                .TextMatrix(lngCurRow, Col_医生) = Nvl(rsData!医生姓名)
                If m_DataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_开始时间) = Format(Nvl(rsData!开始时间), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_终止时间) = Format(Nvl(rsData!终止时间), "yyyy-mm-dd hh:mm:ss")
                End If
                blnAddRow = True
            End If
            
            If blnFind Then
                '排班规则:1-星期排班;2-单日排班;3-双日排班;4-月内轮循;5-轮循不限制;6-特定日期
                '预约控制：0-不作预约限制;1-该号码禁止预约;2-仅禁止三方机构平台的预约
                If Nvl(rsData!上班时段) <> "" Then
                    Select Case m_DataStyle
                    Case Data_Templet
                        If Nvl(rsData!排班规则) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!预约控制) = 1, "-", _
                                                                        IIf(Val(Nvl(rsData!限约数)) = 0, "∞", Nvl(rsData!限约数)))
                        ElseIf Nvl(rsData!排班规则) = 6 Then
                            If .TextMatrix(lngCurRow, lngCurCol) <> "" Then
                                .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & "," & Val(Nvl(rsData!限制项目))
                            Else
                                .TextMatrix(lngCurRow, lngCurCol) = Val(Nvl(rsData!限制项目))
                                .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!上班时段)
                                .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                                .TextMatrix(lngCurRow, lngCurCol + 3) = IIf(Nvl(rsData!预约控制) = 1, "-", _
                                                                            IIf(Val(Nvl(rsData!限约数)) = 0, "∞", Nvl(rsData!限约数)))
                            End If
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!限制项目)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!上班时段)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = IIf(Nvl(rsData!预约控制) = 1, "-", _
                                                                        IIf(Val(Nvl(rsData!限约数)) = 0, "∞", Nvl(rsData!限约数)))
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!限号数)) = 0, "∞", Nvl(rsData!限号数))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!预约控制) = 1, "-", _
                                                                    IIf(Val(Nvl(rsData!限约数)) = 0, "∞", Nvl(rsData!限约数)))
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!上班时段)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!记录ID)
                        If Val(Nvl(rsData!是否临时出诊)) = 1 Then '临时出诊号用蓝色字体显示
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbBlue
                        End If
                        If Val(Nvl(rsData!是否锁定)) = 1 Then '锁定号显示锁的图标
                            .Cell(flexcpPicture, lngCurRow, lngCurCol) = imgLock.Picture
                            .Cell(flexcpPictureAlignment, lngCurRow, lngCurCol) = flexAlignRightBottom
                        End If
                        If Nvl(rsData!停诊开始时间) <> "" Then '停诊号用红色字体显示
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbRed
                        End If
                        If Nvl(rsData!替诊医生姓名) <> "" Then '替诊号用蓝色字体显示并显示替诊医生
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbBlue
'                            .Cell(flexcpPicture, lngCurRow, lngCurCol) = imgReplace.Picture
'                            .Cell(flexcpPictureAlignment, lngCurRow, lngCurCol) = flexAlignRightBottom
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & Chr(10) & Space(4 - Len(Nvl(rsData!替诊医生姓名))) & Nvl(rsData!替诊医生姓名)
                        End If
                        .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!已挂数) & "/" & IIf(Nvl(rsData!限号数) = "", "∞", Nvl(rsData!限号数))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!预约控制) = 1, "-", _
                                                                    Nvl(rsData!已约数) & "/" & IIf(Nvl(rsData!限约数) = "", "∞", Nvl(rsData!限约数)))
                    End Select
                End If
                    
                If blnAddRow Then lngCurRow = lngCurRow + 1
            End If
            rsData.MoveNext
        Loop
        If .Rows > 1 Then '缺省定位行
            .Row = -1 '保证在选择行不变的情况下也触发RowColChange事件
            .Row = IIf(lngOldRow = 0 Or lngOldRow > .Rows - 1, .FixedRows, lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
        End If
        Call SetGridFormat
    End With
    LoadDataByRecordset = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetGridFormat()
    Dim i As Long, j As Long
    Dim lngCurRowInGroup As Long, strSpace As String
    
    With vsfRegistPlan
        '特殊处理，以便能够合并行列
        lngCurRowInGroup = 0 '纵向组内行号
        For i = 2 To .Rows - 1
            If .RowData(i) = -1 Then lngCurRowInGroup = 0
            For j = 0 To .Cols - 1
                If .RowData(i) = -1 Then Exit For
                If .TextMatrix(i, j) = "" Then .TextMatrix(i, j) = " " '防止内容为空不能合并
                If .RowData(i - 1) <> -1 And j >= .FixedCols Then '是否为星期数据
                    If (j - .FixedCols) Mod 3 = 0 Then '"时间段"列
                        If .TextMatrix(i, j) = " " Then '合并后面的空行
                            .TextMatrix(i, j) = .TextMatrix(i - 1, j)
                            .TextMatrix(i, j + 1) = .TextMatrix(i - 1, j + 1)
                            .TextMatrix(i, j + 2) = .TextMatrix(i - 1, j + 2)
                        Else
                            strSpace = Space(lngCurRowInGroup Mod 2) '填充空格，防止内容相同合并
                            .TextMatrix(i, j + 1) = strSpace & .TextMatrix(i, j + 1) & strSpace
                            .TextMatrix(i, j + 2) = strSpace & .TextMatrix(i, j + 2) & strSpace
                        End If
                        '模板非星期排班
                        If m_DataStyle = Data_Templet And j = .Cols - 4 Then
                            If .TextMatrix(i, j) = " " Then '合并后面的空行
                                .TextMatrix(i, j + 3) = .TextMatrix(i - 1, j + 3)
                            Else
                                strSpace = Space(lngCurRowInGroup Mod 2) '填充空格，防止内容相同合并
                                .TextMatrix(i, j + 3) = strSpace & .TextMatrix(i, j + 3) & strSpace
                            End If
                            j = j + 1
                        End If
                        j = j + 2
                    End If
                End If
            Next
            If .RowData(i) <> -1 Then lngCurRowInGroup = lngCurRowInGroup + 1
        Next
        
        '行背景色
        For i = 2 To .Rows - 1
            .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = .RowData(i)
        Next
        
        '字体颜色
        For i = 2 To .Rows - 1 '无安排的号源用灰色显示
            If Val(.TextMatrix(i, COL_安排ID)) = 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayText
            End If
        Next
        If m_DataStyle = Data_Plan Then
            For j = .FixedCols To .Cols - 1 Step 3 '失效的安排用灰色显示
                If CDate(.Cell(flexcpData, 0, j)) < Format(Now, "yyyy-mm-dd") Then
                    .Cell(flexcpForeColor, 0, j, .Rows - 1, j + 2) = vbGrayText
                End If
            Next
        End If
            
        
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
        If .RowData(.Row) = -1 And .Rows > .Row + 1 Then .Row = .Row + 1
        If .Col < .FixedCols And .Cols > .FixedCols Then .Col = .FixedCols
        Call SetVSFSelectedRangeColor(.Row, .Col, False)
    End With
End Sub

Private Sub SetVSFSelectedRangeColor(ByVal Row As Long, ByVal Col As Long, ByVal blnOld As Boolean)
    '功能：设置选择行颜色,.RowData中存了颜色值
    Dim i As Long
    Dim lngRowStart As Long, lngRowEnd As Long '起始行和终止行
    Dim lngColStart As Long, lngColEnd As Long '起始列和终止列

    On Error Resume Next
    With vsfRegistPlan
        If Not .Visible Then Exit Sub
        If Row > .FixedRows - 1 And Col > .FixedCols - 1 Then
'            lngRowStart = .FixedRows
'            For i = Row To .FixedRows Step -1
'                If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
'            Next
'            lngRowEnd = .Rows - 1
'            For i = Row + 1 To .Rows - 1
'                If .RowData(i) = -1 Then lngRowEnd = i - 1: Exit For
'            Next
            lngRowStart = Row
            lngRowEnd = Row
            lngColStart = GetItemNameCol(Col, .FixedCols) '确定"时间段"列
            If m_DataStyle = Data_Templet And lngColStart = .Cols - 4 Then lngColStart = lngColStart + 1
            lngColEnd = lngColStart + 2
            .Cell(flexcpBackColor, lngRowStart, lngColStart, lngRowEnd, lngColEnd) = IIf(blnOld, .RowData(lngRowStart), .BackColorSel)
        End If
    End With
End Sub

Private Function GetItemNameCol(ByVal lngCurCol As Long, ByVal lngFixedCols As Long) As Long
    '确定"时段"列的列索引
    GetItemNameCol = lngCurCol - Choose(((lngCurCol - lngFixedCols) Mod 3) + 1, 0, 1, 2)
End Function

Private Sub GetGroupRange(ByVal lngCurRow As Long, ByRef lngRowStart As Long, ByRef lngRowEnd As Long)
    '当前组的行索引范围
    Dim i As Integer
    
    With vsfRegistPlan
        lngRowStart = .FixedRows
        For i = lngCurRow To .FixedRows Step -1
            If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
        Next
        lngRowEnd = .Rows - 1
        For i = lngCurRow + 1 To .Rows - 1
            If .RowData(i) = -1 And i <> .Rows - 1 Then lngRowEnd = i - 1: Exit For
        Next
    End With
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetVSFSelectedRangeColor(OldRow, OldCol, True)
    Call SetVSFSelectedRangeColor(NewRow, NewCol, False)
    RaiseEvent AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfRegistPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COl_空列 Then Cancel = True
End Sub

Private Sub vsfRegistPlan_DblClick()
    Dim i As Long, strSort As String
    Dim lngCol As Long, lngRow As Long
    
    RaiseEvent DblClick
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If (lngRow = 0 Or lngRow = 1) And Not mrsData Is Nothing Then
        '按列排序
        Select Case lngCol
        Case COL_号类
            strSort = SortCircle(lngCol, "号类") & "号码,"
        Case col_号码
            strSort = SortCircle(lngCol, "号码")
        Case COL_科室
            strSort = SortCircle(lngCol, "科室") & "号码,"
        Case COL_项目
            strSort = SortCircle(lngCol, "收费项目") & "号码,"
        Case Col_医生
            strSort = SortCircle(lngCol, "医生姓名") & "号码,"
        End Select
        If strSort <> "" Then
            Select Case m_DataStyle
            Case Data_FixedRule
                strSort = "安排ID," & strSort & "开始时间,终止时间,限制项目,上班时段"
            Case Data_Templet
                strSort = "安排ID," & strSort & "限制项目,上班时段"
            Case Data_Plan
                strSort = strSort & "出诊日期,上班时段"
            End Select
            mrsData.Sort = strSort
            Call LoadDataByRecordset(mrsData)
        End If
    End If
End Sub

Private Function SortCircle(ByVal lngCol As Long, ByVal strColName As String) As String
    'Cell(flexcpData, 1, lngCol)记录了当前排序方式，注意在重新加载数据时清除
    Select Case vsfRegistPlan.Cell(flexcpData, 1, lngCol)
    Case ""
        If lngCol = col_号码 Then '号码列初始时就是以升序排列的
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "DESC"
            SortCircle = strColName & " DESC,"
        Else
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        End If
    Case "ASC" '升序
        vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "DESC"
        SortCircle = strColName & " Desc,"
    Case "DESC" '降序
        If lngCol = col_号码 Then '号码列要么升序要么降序
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        Else
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "-"
            SortCircle = ""
        End If
    Case "-" '不排序
        vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
        SortCircle = strColName & " Asc,"
    End Select
End Function

Private Sub vsfRegistPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub vsfRegistPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim varData As Variant, strTemp As String
    Dim lngRow As Long, lngCol As Long
    Dim sngLeft As Single, sngTop As Single

    On Error Resume Next
    If m_DataStyle <> Data_Plan Then Exit Sub
    With vsfRegistPlan
        If Not .Visible Then Exit Sub
        
        lngRow = .MouseRow: lngCol = .MouseCol
        If .Tag = lngRow & "," & lngCol Then Exit Sub
        .Tag = lngRow & "," & lngCol
        picToolTip.Visible = False
        
        If lngRow < .FixedRows Or lngCol < .FixedCols Then Exit Sub
        If (lngCol - .FixedCols) Mod 3 = 0 Then Exit Sub '"时段"列退出
        strTemp = .TextMatrix(lngRow, lngCol)
        If (strTemp = "" Or InStr(strTemp, "/") = 0) And strTemp <> "-" Then Exit Sub
        
        '2.显示内容
        If strTemp = "-" Then
            lblToolTip1.Caption = "禁止预约！"
            lblToolTip2.Visible = False
        Else
            lblToolTip2.Visible = True
            varData = Split(strTemp, "/")
            If (lngCol - .FixedCols) Mod 3 = 1 Then
                lblToolTip1.Caption = "已挂数：" & Trim(varData(0))
                lblToolTip2.Caption = "限号数：" & IIf(Trim(varData(1)) = "∞", "不限制", Trim(varData(1)))
            ElseIf (lngCol - .FixedCols) Mod 3 = 2 Then
                lblToolTip1.Caption = "已约数：" & Trim(varData(0))
                lblToolTip2.Caption = "限约数：" & IIf(Trim(varData(1)) = "∞", "不限制", Trim(varData(1)))
            End If
        End If
        
        '3.提示框的位置
        sngLeft = .Cell(flexcpLeft, lngRow, lngCol) + .ColWidth(lngCol) - 50
        sngTop = .Cell(flexcpTop, lngRow, lngCol) + .RowHeight(lngRow) - 10
        If sngLeft + picToolTip.Width > .Width Then
            sngLeft = .Cell(flexcpLeft, lngRow, lngCol) - picToolTip.Width + 50
        End If
        If sngTop + picToolTip.Height > .Height Then
            sngTop = .Cell(flexcpTop, lngRow, lngCol) - picToolTip.Height + 20
        End If
        picToolTip.Move sngLeft, sngTop
        picToolTip.Visible = True
    End With
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(UserControl.Parent, UserControl.Parent.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
End Sub
