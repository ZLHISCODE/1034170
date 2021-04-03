VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品批量选择"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11865
   Icon            =   "frmBatchSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12.753
   ScaleMode       =   0  'User
   ScaleWidth      =   20.929
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDrug 
      Height          =   5535
      Left            =   1560
      ScaleHeight     =   5475
      ScaleWidth      =   3915
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
      Begin VSFlex8Ctl.VSFlexGrid vsfDrug 
         Height          =   5470
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3930
         _cx             =   6932
         _cy             =   9648
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBatchSelect.frx":000C
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
   Begin VB.CommandButton cmdQuit 
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   10320
      TabIndex        =   6
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   9240
      TabIndex        =   5
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton cmdCal 
      Caption         =   "清空(&O)"
      Height          =   300
      Left            =   8160
      TabIndex        =   4
      Top             =   6720
      Width           =   975
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   960
      TabIndex        =   3
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox txtSelect 
      Height          =   300
      Left            =   1560
      TabIndex        =   1
      Top             =   180
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   10920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":0081
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":061B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchSelect.frx":6E7D
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfSelectDrug 
      Height          =   5925
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   11775
      _cx             =   20770
      _cy             =   10451
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBatchSelect.frx":7417
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
      ExplorerBar     =   1
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
      VirtualData     =   0   'False
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
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "查找"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   6780
      Width           =   360
   End
   Begin VB.Label lblCalss 
      AutoSize        =   -1  'True
      Caption         =   "请输入品种简码"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmBatchSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintUnit As Integer '本模块中设置的显示单位 0-药库单位;1-门诊单位;2-住院单位;3-售价单位
Private Const mlngRowHeight As Long = 300 '表格中各行行高
Private mrsReturn As ADODB.Recordset        '返回选定药品数据
Private mblnOk As Boolean   '记录是否是点击的确定按钮
Private mrsFindName As ADODB.Recordset '记录查询数据集
Private mstrMatch  As String '0-双向匹配 1-单向右匹配

'各单位
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
Private Const MStrCaption As String = "药品批量选择"

Private Enum vsfSelectDrugCol
    药品ID = 0
    药品信息 = 1
    药品编码
    商品名
    通用名
    规格
    产地
    单位
    售价单位
    门诊单位
    门诊系数
    住院单位
    住院系数
    药库单位
    药库系数
    类型
    售价
    成本价
    指导批价
    指导售价
    总列数
End Enum

Public Sub showMe(ByVal frmParent As Form, ByRef rsTemp As ADODB.Recordset, ByRef blnOk As Boolean)
    Me.Show vbModal, frmParent
    blnOk = mblnOk
    Set rsTemp = mrsReturn
End Sub

Private Sub initVsflexgrid()
    With vsfSelectDrug
        .Editable = flexEDNone
        .Cols = vsfSelectDrugCol.总列数
        .rows = 1
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '不能多选
        .SelectionMode = flexSelectionByRow '整行选择
        .ExplorerBar = flexExMove '移动
        .AllowUserResizing = flexResizeBoth  '可以改变行列宽度

        '设置列宽
        .ColWidth(vsfSelectDrugCol.药品ID) = 0
        .ColWidth(vsfSelectDrugCol.药品信息) = 3000
        .ColWidth(vsfSelectDrugCol.药品编码) = 0
        .ColWidth(vsfSelectDrugCol.商品名) = 0
        .ColWidth(vsfSelectDrugCol.通用名) = 0
        .ColWidth(vsfSelectDrugCol.产地) = 1500
        .ColWidth(vsfSelectDrugCol.单位) = 800
        
        .ColWidth(vsfSelectDrugCol.售价单位) = 0
        .ColWidth(vsfSelectDrugCol.门诊单位) = 0
        .ColWidth(vsfSelectDrugCol.门诊系数) = 0
        .ColWidth(vsfSelectDrugCol.住院单位) = 0
        .ColWidth(vsfSelectDrugCol.住院系数) = 0
        .ColWidth(vsfSelectDrugCol.药库单位) = 0
        .ColWidth(vsfSelectDrugCol.药库系数) = 0
        
        .ColWidth(vsfSelectDrugCol.类型) = 1000
        .ColWidth(vsfSelectDrugCol.售价) = 1500
        .ColWidth(vsfSelectDrugCol.成本价) = 1500
        .ColWidth(vsfSelectDrugCol.指导批价) = 1500
        .ColWidth(vsfSelectDrugCol.指导售价) = 1500
        '设置列头
        .TextMatrix(0, vsfSelectDrugCol.药品ID) = "药品id"
        .TextMatrix(0, vsfSelectDrugCol.药品信息) = "药品信息"
        .TextMatrix(0, vsfSelectDrugCol.药品编码) = "药品编码"
        .TextMatrix(0, vsfSelectDrugCol.商品名) = "商品名"
        .TextMatrix(0, vsfSelectDrugCol.通用名) = "通用名"
        .TextMatrix(0, vsfSelectDrugCol.规格) = "规格"
        .TextMatrix(0, vsfSelectDrugCol.产地) = "产地"
        .TextMatrix(0, vsfSelectDrugCol.单位) = "单位"
        
        .TextMatrix(0, vsfSelectDrugCol.售价单位) = "售价单位"
        .TextMatrix(0, vsfSelectDrugCol.门诊单位) = "门诊单位"
        .TextMatrix(0, vsfSelectDrugCol.门诊系数) = "门诊系数"
        .TextMatrix(0, vsfSelectDrugCol.住院单位) = "住院单位"
        .TextMatrix(0, vsfSelectDrugCol.住院系数) = "住院系数"
        .TextMatrix(0, vsfSelectDrugCol.药库单位) = "药库单位"
        .TextMatrix(0, vsfSelectDrugCol.药库系数) = "药库系数"
        
        .TextMatrix(0, vsfSelectDrugCol.类型) = "类型"
        .TextMatrix(0, vsfSelectDrugCol.售价) = "售价"
        .TextMatrix(0, vsfSelectDrugCol.成本价) = "成本价"
        .TextMatrix(0, vsfSelectDrugCol.指导批价) = "指导批价"
        .TextMatrix(0, vsfSelectDrugCol.指导售价) = "指导售价"

        .ColAlignment(vsfSelectDrugCol.药品ID) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.药品信息) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.药品编码) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.规格) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.产地) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.单位) = flexAlignCenterCenter
        .ColAlignment(vsfSelectDrugCol.类型) = flexAlignLeftCenter
        .ColAlignment(vsfSelectDrugCol.售价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.成本价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导批价) = flexAlignRightCenter
        .ColAlignment(vsfSelectDrugCol.指导售价) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

'
'Private Sub setTvwInfo()
'    '为树表填充数据
'    Dim objNode As Node
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'
'    gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
'              " Where Instr([1],编码,1) > 0 " & _
'              " Order by 编码"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, mstrCaption, "567")
'
'    If rsTemp Is Nothing Then
'        Exit Sub
'    End If
'
'    With tvwDrug
'        .Nodes.Clear
'        Do While Not rsTemp.EOF
'            .Nodes.Add , , "Root" & rsTemp!名称, rsTemp!名称, 1, 1
'            .Nodes("Root" & rsTemp!名称).Tag = rsTemp!编码
'            rsTemp.MoveNext
'        Loop
'    End With
'
'
'    gstrSQL = "Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
'                " From 诊疗分类目录" & _
'                " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                " Start With 上级id Is Null" & _
'                " Connect By Prior ID = 上级id"
'
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "分类查询")
'    With rsTemp
'        Do While Not .EOF
'           If IsNull(!上级id) Then
'                Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'            Else
'                Set objNode = tvwDrug.Nodes.Add("K_" & !上级id, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'            End If
'            objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'            .MoveNext
'        Loop
'    End With
'
'    If optVariety.Value = True Then
'        gstrSQL = "Select ID, 分类id, 编码, 名称, Decode(类别, 5, '西成药', 6, '中成药', 7, '中草药') 分类, '品种' As 类别" & _
'                  "  From 诊疗项目目录" & _
'                  "  Where 分类id In (Select ID" & _
'                                   " From 诊疗分类目录" & _
'                                   " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                                   " Start With 上级id Is Null" & _
'                                   " Connect By Prior ID = 上级id)"
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "品种")
'
'        With rsTemp
'            Do While Not .EOF
'                Set objNode = tvwDrug.Nodes.Add("K_" & !分类id, 4, !类别 & "K_" & !Id, !名称 & "-品种", 1, 1)
'                objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                .MoveNext
'            Loop
'        End With
'    End If
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

'Private Sub cmdSelect_Click()
'    picDrug.Visible = True
'    tvwDrug.Visible = True
'    Call setTvwInfo
'End Sub

Private Sub cmdCal_Click()
    With vsfSelectDrug
        If MsgBox("确定要清空所有已经选择的药品？", vbYesNo, gstrSysName) = vbYes Then
            .rows = 1
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    Dim intRow As Integer
    Set mrsReturn = New ADODB.Recordset

    With mrsReturn
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品编码", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "时价", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
        
        .Fields.Append "售价单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "门诊包装", adDouble, 11, adFldIsNullable
        .Fields.Append "住院单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "住院包装", adDouble, 11, adFldIsNullable
        .Fields.Append "药库单位", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药库包装", adDouble, 11, adFldIsNullable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With

    With vsfSelectDrug
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, vsfSelectDrugCol.药品ID) = "" Then Exit For
            mrsReturn.AddNew
            mrsReturn!药品ID = .TextMatrix(intRow, vsfSelectDrugCol.药品ID)
            mrsReturn!药品编码 = .TextMatrix(intRow, vsfSelectDrugCol.药品编码)
            mrsReturn!商品名 = .TextMatrix(intRow, vsfSelectDrugCol.商品名)
            mrsReturn!通用名 = .TextMatrix(intRow, vsfSelectDrugCol.通用名)
            mrsReturn!规格 = .TextMatrix(intRow, vsfSelectDrugCol.规格)
            mrsReturn!时价 = IIf(.TextMatrix(intRow, vsfSelectDrugCol.类型) = "时价", 1, 0)
            mrsReturn!产地 = .TextMatrix(intRow, vsfSelectDrugCol.产地)
            mrsReturn!售价单位 = .TextMatrix(intRow, vsfSelectDrugCol.售价单位)
            mrsReturn!门诊单位 = .TextMatrix(intRow, vsfSelectDrugCol.门诊单位)
            mrsReturn!门诊包装 = .TextMatrix(intRow, vsfSelectDrugCol.门诊系数)
            mrsReturn!住院单位 = .TextMatrix(intRow, vsfSelectDrugCol.住院单位)
            mrsReturn!住院包装 = .TextMatrix(intRow, vsfSelectDrugCol.住院系数)
            mrsReturn!药库单位 = .TextMatrix(intRow, vsfSelectDrugCol.药库单位)
            mrsReturn!药库包装 = .TextMatrix(intRow, vsfSelectDrugCol.药库系数)
            
            mrsReturn.Update
        Next
    End With
    mblnOk = True
    
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdSelect_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDrug.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim intUnitTemp As Integer
    '获取设置的单位
    mintUnit = Val(zlDatabase.GetPara("药品单位", glngSys, 1333, 1))
    Select Case mintUnit
        Case 0 '药库
            intUnitTemp = 4
        Case 1 '住院
            intUnitTemp = 3
        Case 2 '门诊
            intUnitTemp = 2
        Case 3 '售价
            intUnitTemp = 1
    End Select
    '获取各级单位精度
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    
    mstrMatch = IIf(zlDatabase.GetPara("输入匹配", , , 0) = "0", "%", "")
    mblnOk = False
    Call initVsflexgrid
    
    Call RestoreWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, MStrCaption)
End Sub

Private Sub optClass_Click()
    picDrug.Visible = False
    lblCalss.Caption = "分类"
End Sub

Private Sub optClassSub_Click()
    picDrug.Visible = False
    lblCalss.Caption = "分类(含子类)"
End Sub

Private Sub optVariety_Click()
    picDrug.Visible = False
    lblCalss.Caption = "品种"
End Sub

'Private Sub tvwDrug_NodeClick(ByVal Node As MSComctlLib.Node)
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo errHandle
'    If Node.Key Like "Root" Then Exit Sub
'
'    gstrSQL = "select id,编码,名称,计算单位 from 诊疗项目目录 where  Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' and 分类id=[1]"
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查询品种", Mid(Node.Key, InStr(1, Node.Key, "_") + 1))
'
'    Set vsfDetails.DataSource = rsTemp
'
'    Exit Sub
'errHandle:
'    If errcenter() = 1 Then Resume
'    Call saveerrlog
'End Sub

'Private Sub tvwDrug_DblClick()
'    '用来向界面中传入值
'    Dim lngId As Long
'    Dim rsTemp As ADODB.Recordset
'    Dim intRow As Integer
'    Dim i As Integer
'    Dim blnDou As Boolean '重复数据
'    Dim dbl换算系数 As Double
'    Dim strUnit As String   '单位
'
'    On Error GoTo errHandle
'    With tvwDrug
'        If optVariety.Value = True Then
'            If InStr(1, .SelectedItem.Text, "-品种") <= 0 Then
'                Exit Sub
'            End If
'            gstrSQL = "Select Distinct a.药品id, c.编码 As 药品编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位 As 售价单位, a.门诊单位, a.门诊包装," & _
'                                        " a.住院单位 , a.住院包装, a.药库单位, a.药库包装, a.成本价, e.现价, a.指导批发价, a.指导零售价" & _
'                        " From 药品规格 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D,收费价目 E" & _
'                        " Where a.药名id = b.Id And a.药品id = c.Id And c.Id = d.收费细目id(+) and a.药品id=e.收费细目id and sysdate between e.执行日期 and e.终止日期 and b.id=[1] order by c.编码"
'        Else
'            If InStr(1, .SelectedItem.Text, "-分类") <= 0 Then
'                Exit Sub
'            End If
'                If optClassSub.Value = True Then '分类下子节点
'                    gstrSQL = "(Select ID From 诊疗分类目录 Where 类型 In (1, 2, 3) Start With ID = [1] Connect By Prior ID = 上级id) A,"
'                Else '本分类
'                    gstrSQL = "(select id from 诊疗分类目录 where 类型 in (1,2,3) and id=[1]) A,"
'                End If
'
'                gstrSQL = "Select Distinct c.药品id, d.编码 As 药品编码, d.名称 As 通用名, f.商品名, d.规格, d.是否变价 As 时价, d.产地, d.计算单位 As 售价单位, c.门诊单位, c.门诊包装," & _
'                                        " c.住院单位 , c.住院包装, c.药库单位, c.药库包装, c.成本价, e.现价, c.指导批发价, c.指导零售价 " & _
'                        " From " & gstrSQL & " 诊疗项目目录 B, 药品规格 C," & _
'                             " 收费项目目录 D, 收费价目 E, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) F" & _
'                        " Where a.Id = b.分类id And b.Id = c.药名id And c.药品id = d.Id And d.Id = e.收费细目id And e.收费细目id = f.收费细目id(+) And" & _
'                              " Sysdate Between e.执行日期 And e.终止日期 order by d.编码"
'        End If
'        lngId = Mid(.SelectedItem.Key, InStr(1, .SelectedItem.Key, "K_") + 2)
'
'        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查询药品", lngId)
'        If rsTemp.RecordCount = 0 Then
'            Exit Sub
'        End If
'    End With
'
'    With vsfSelectDrug
'        For intRow = 0 To rsTemp.RecordCount - 1
'            blnDou = False
'            For i = 1 To .rows - 1
'                If .TextMatrix(i, vsfSelectDrugCol.药品id) = rsTemp!药品id Then
'                    blnDou = True
'                End If
'            Next
'            If blnDou = False Then
'                .rows = .rows + 1
'                .RowHeight(.rows - 1) = mlngRowHeight
'
'                Select Case mintUnit
'                    Case 0
'                        dbl换算系数 = rsTemp!药库包装
'                        strUnit = rsTemp!药库单位
'                    Case 1
'                        dbl换算系数 = rsTemp!住院包装
'                        strUnit = rsTemp!住院单位
'                    Case 2
'                        dbl换算系数 = rsTemp!门诊包装
'                        strUnit = rsTemp!门诊单位
'                    Case 3
'                        dbl换算系数 = 1
'                        strUnit = rsTemp!售价单位
'                End Select
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品id) = rsTemp!药品id
'                If gint药品名称显示 = 1 Then
'                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
'                Else
'                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & rsTemp!通用名
'                End If
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品编码) = rsTemp!药品编码
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.商品名) = IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.通用名) = IIf(IsNull(rsTemp!通用名), "", rsTemp!通用名)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.单位) = strUnit
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价单位) = rsTemp!售价单位
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊单位) = rsTemp!门诊单位
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊系数) = rsTemp!门诊包装
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院单位) = rsTemp!住院单位
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院系数) = rsTemp!住院包装
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库单位) = rsTemp!药库单位
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库系数) = rsTemp!药库包装
'
'
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.类型) = IIf(rsTemp!时价 = 1, "时价", "定价")
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价) = GetFormat(dbl换算系数 * rsTemp!现价, mintPriceDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.成本价) = GetFormat(dbl换算系数 * rsTemp!成本价, mintCostDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导批价) = GetFormat(dbl换算系数 * rsTemp!指导批发价, mintCostDigit)
'                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导售价) = GetFormat(dbl换算系数 * rsTemp!指导零售价, mintPriceDigit)
'
'            End If
'            rsTemp.MoveNext
'        Next
'        picDrug.Visible = False
'    End With
'
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub
    
    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim str药名 As String
    Dim lngRow As Long
    
    '查找药品
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] ) " & _
                  "Order By 药品编码 "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '如果到底了，则返回第1条记录
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = mrsFindName!药品编码 & mrsFindName!通用名
        Else
            str药名 = mrsFindName!药品编码 & IIf(IsNull(mrsFindName!商品名), mrsFindName!通用名, mrsFindName!商品名)
        End If
        
        For lngRow = 1 To vsfSelectDrug.rows - 1
            lngFindRow = vsfSelectDrug.FindRow(str药名, lngRow, CLng(vsfSelectDrugCol.药品信息), True, True)
            If lngFindRow > 0 Then
                vsfSelectDrug.Select lngFindRow, 1, lngFindRow, vsfSelectDrug.Cols - 1
                vsfSelectDrug.TopRow = lngFindRow
                Exit For
            End If
        Next
        
        If lngFindRow > 0 Then  '查询到数据后就移动下下一条并退出本次查询
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext '未查询到数据则移动到下一条数据集继续查询
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtSelect_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub

Private Sub txtSelect_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim rsPinzhong As ADODB.Recordset
    Dim objNode As Node
    Dim lng分类id As Long
    Dim i As Integer
    
    If KeyCode = vbKeyReturn Then
    
        On Error GoTo errHandle
        
        If Trim(txtSelect.Text) = "" Then Exit Sub
                
        gstrSQL = "Select Distinct a.id,a.编码,a.名称" & _
                  "  From 诊疗项目目录 A, 诊疗项目别名 B" & _
                    " Where a.Id = b.诊疗项目id(+) And a.类别 In ('5', '6', '7') And Sysdate Between 建档时间 And 撤档时间 And" & _
                         " (a.编码 Like [1] Or a.名称 Like [1] Or b.名称 Like [1] Or b.简码 Like [1])"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
        If rsTemp.RecordCount = 0 Then
            MsgBox "未查询到品种！", vbInformation, gstrSysName
            txtSelect.SetFocus
            txtSelect.SelStart = 1
            txtSelect.SelLength = Len(txtSelect.Text)
        Else
            picDrug.Visible = True
            vsfDrug.Visible = True
            Set vsfDrug.DataSource = rsTemp
            vsfDrug.SetFocus
            vsfDrug.Row = 1
        End If
        With vsfDrug
            For i = 0 To .rows - 1
                .RowHeight(i) = mlngRowHeight
            Next
        End With
'        gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
'                  " Where Instr([1],编码,1) > 0 " & _
'                  " Order by 编码"
'        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, "567")
'
'        If rsTemp Is Nothing Then
'            Exit Sub
'        End If
'
'        With tvwDrug
'            .Nodes.Clear
'            Do While Not rsTemp.EOF
'                .Nodes.Add , , "Root" & rsTemp!名称, rsTemp!名称, 1, 1
'                .Nodes("Root" & rsTemp!名称).Tag = rsTemp!编码
'                rsTemp.MoveNext
'            Loop
'        End With
        
'        If optVariety.Value = True Then '品种被选中
'            gstrSQL = "Select a.Id, a.上级id, a.编码, a.名称, Decode(a.类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
'                        " From 诊疗分类目录 A," & _
'                             " (Select Distinct a.分类id" & _
'                               " From 诊疗项目目录 A, 诊疗项目别名 B" & _
'                               " Where a.Id = b.诊疗项目id(+) And a.类别 In ('5', '6', '7') And Sysdate Between 建档时间 And 撤档时间 And" & _
'                                     " (a.编码 Like [1] Or a.名称 Like [1] Or b.名称 Like [1] Or b.简码 Like [1])) B" & _
'                        " Where a.Id = b.分类id And Nvl(To_Char(a.撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'                        " Start With a.上级id Is Null" & _
'                        " Connect By Prior a.Id = a.上级id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!上级id) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !上级id, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    End If
'                    objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                    .MoveNext
'                Loop
'
'                rsTemp.MoveFirst
'                Do While Not rsTemp.EOF
'                    lng分类id = rsTemp!Id
'                    gstrSQL = "Select Distinct a.Id, a.分类id, a.编码, a.名称, Decode(a.类别, '5', '西成药', '6', '中成药', '7', '中草药') 分类, '品种' As 类别" & _
'                                " From 诊疗项目目录 A" & _
'                                " Where a.类别 In ('5', '6', '7') And a.分类id=[1] and Sysdate Between a.建档时间 And a.撤档时间"
'                    Set rsPinzhong = zlDatabase.OpenSQLRecord(gstrSQL, "品种", lng分类id)
'
'                    Do While Not rsPinzhong.EOF
'                        Set objNode = tvwDrug.Nodes.Add("K_" & rsPinzhong!分类id, 4, rsPinzhong!类别 & "K_" & rsPinzhong!Id, rsPinzhong!名称 & "-品种", 1, 1)
'                        objNode.Tag = rsPinzhong!分类 & "-" & rsPinzhong!类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                        rsPinzhong.MoveNext
'                    Loop
'                    rsTemp.MoveNext
'                Loop
'            End With
'        Else
'            gstrSQL = "Select ID, 上级id, 编码, 名称, 分类, 类别　 from ( " & _
'                    "Select Distinct ID, 上级id, 编码, 名称, 分类, 类别" & _
'                        " From (Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
'                               " From 诊疗分类目录" & _
'                               " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
'                                     " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1])" & _
'                               " Start With 上级id Is Null" & _
'                               " Connect By Prior ID = 上级id" & _
'                               " Union All" & _
'                               " Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & _
'                               " From 诊疗分类目录" & _
'                               " Where ID In (Select 上级id" & _
'                                            " From 诊疗分类目录" & _
'                                            " Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' And" & _
'                                                  " (编码 Like [1] Or 名称 Like [1] Or 简码 Like [1]))))" & _
'                        " Start With 上级id Is Null" & _
'                        " Connect By Prior ID = 上级id"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询品种", "%" & UCase(txtSelect.Text) & mstrMatch)
'            If rsTemp.RecordCount = 0 Then Exit Sub
'
'            With rsTemp
'                Do While Not .EOF
'                   If IsNull(!上级id) Then
'                        Set objNode = tvwDrug.Nodes.Add("Root" & !分类, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    Else
'                        Set objNode = tvwDrug.Nodes.Add("K_" & !上级id, 4, "K_" & !Id, !名称 & "-分类", 1, 1)
'                    End If
'                    objNode.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
'                    .MoveNext
'                Loop
'            End With
'        End If
'        tvwDrug.SetFocus
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_DblClick()
    Dim lngId As Long
    Dim rsTemp As ADODB.Recordset
    Dim intRow As Integer
    Dim i As Integer
    Dim blnDou As Boolean '重复数据
    Dim dbl换算系数 As Double
    Dim strUnit As String   '单位

    On Error GoTo errHandle
    With vsfDrug
        If Val(.TextMatrix(.Row, 0)) = 0 Then
            Exit Sub
        End If
        gstrSQL = "Select Distinct a.药品id, c.编码 As 药品编码, c.名称 As 通用名, d.商品名, c.规格, c.是否变价 As 时价, c.产地, c.计算单位 As 售价单位, a.门诊单位, a.门诊包装," & _
                                    " a.住院单位 , a.住院包装, a.药库单位, a.药库包装, a.成本价, e.现价, a.指导批发价, a.指导零售价" & _
                    " From 药品规格 A, 诊疗项目目录 B, 收费项目目录 C, (Select 名称 As 商品名, 收费细目id From 收费项目别名 Where 性质 = 3) D,收费价目 E" & _
                    " Where a.药名id = b.Id And a.药品id = c.Id And c.Id = d.收费细目id(+) and a.药品id=e.收费细目id and sysdate between e.执行日期 and e.终止日期 and b.id=[1] " & _
                    " And (c.撤档时间 = to_date('3000-01-01','yyyy-mm-dd') or c.撤档时间 is null ) order by c.编码"
        
        lngId = Val(.TextMatrix(.Row, 0))

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询药品", lngId)
        If rsTemp.RecordCount = 0 Then
            Exit Sub
        End If
    End With

    With vsfSelectDrug
        For intRow = 0 To rsTemp.RecordCount - 1
            blnDou = False
            For i = 1 To .rows - 1
                If .TextMatrix(i, vsfSelectDrugCol.药品ID) = rsTemp!药品ID Then
                    blnDou = True
                End If
            Next
            If blnDou = False Then
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
            
                Select Case mintUnit
                    Case 0
                        dbl换算系数 = rsTemp!药库包装
                        strUnit = rsTemp!药库单位
                    Case 1
                        dbl换算系数 = rsTemp!住院包装
                        strUnit = rsTemp!住院单位
                    Case 2
                        dbl换算系数 = rsTemp!门诊包装
                        strUnit = rsTemp!门诊单位
                    Case 3
                        dbl换算系数 = 1
                        strUnit = rsTemp!售价单位
                End Select
                                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品ID) = rsTemp!药品ID
                If gint药品名称显示 = 1 Then
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & IIf(IsNull(rsTemp!商品名), rsTemp!通用名, rsTemp!商品名)
                Else
                    .TextMatrix(.rows - 1, vsfSelectDrugCol.药品信息) = "[" & rsTemp!药品编码 & "]" & rsTemp!通用名
                End If

                .TextMatrix(.rows - 1, vsfSelectDrugCol.药品编码) = rsTemp!药品编码
                .TextMatrix(.rows - 1, vsfSelectDrugCol.商品名) = IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.通用名) = IIf(IsNull(rsTemp!通用名), "", rsTemp!通用名)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.规格) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.产地) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.单位) = strUnit
                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价单位) = rsTemp!售价单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊单位) = rsTemp!门诊单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.门诊系数) = rsTemp!门诊包装
                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院单位) = rsTemp!住院单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.住院系数) = rsTemp!住院包装
                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库单位) = rsTemp!药库单位
                .TextMatrix(.rows - 1, vsfSelectDrugCol.药库系数) = rsTemp!药库包装
                
                
                .TextMatrix(.rows - 1, vsfSelectDrugCol.类型) = IIf(rsTemp!时价 = 1, "时价", "定价")
                .TextMatrix(.rows - 1, vsfSelectDrugCol.售价) = GetFormat(dbl换算系数 * rsTemp!现价, mintPriceDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.成本价) = GetFormat(dbl换算系数 * rsTemp!成本价, mintCostDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导批价) = GetFormat(dbl换算系数 * rsTemp!指导批发价, mintCostDigit)
                .TextMatrix(.rows - 1, vsfSelectDrugCol.指导售价) = GetFormat(dbl换算系数 * rsTemp!指导零售价, mintPriceDigit)
                
            End If
            rsTemp.MoveNext
        Next
        picDrug.Visible = False
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call vsfDrug_DblClick
    End If
End Sub

Private Sub vsfSelectDrug_GotFocus()
    If picDrug.Visible = True Then
        picDrug.Visible = False
    End If
End Sub
